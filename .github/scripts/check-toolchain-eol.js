#!/usr/bin/env node

/**
 * Inventory checked-in Node.js release-line selectors and compare monitored
 * selectors with the official Node.js release schedule.
 */

const fs = require('fs');
const path = require('path');
const { globSync } = require('glob');
const semver = require('semver');
const yaml = require('yaml');

const DEFAULT_NODE_SCHEDULE_URL =
    'https://raw.githubusercontent.com/nodejs/Release/main/schedule.json';
const DEFAULT_WARNING_WINDOW_DAYS = 180;
const MILLIS_PER_DAY = 24 * 60 * 60 * 1000;
const DEFAULT_FETCH_TIMEOUT_MS = 15000;

function toPosixPath(filePath) {
    return filePath.split(path.sep).join('/');
}

function repoRelativePath(repoRoot, absolutePath) {
    return toPosixPath(path.relative(repoRoot, absolutePath));
}

function assertPathWithinRepo(repoRoot, candidatePath) {
    const rootResolved = path.resolve(repoRoot);
    const candidateResolved = path.resolve(candidatePath);
    const relative = path.relative(rootResolved, candidateResolved);
    if (relative === '' || (!relative.startsWith('..') && !path.isAbsolute(relative))) {
        return candidateResolved;
    }
    throw new Error(`path escapes repository root: ${candidatePath}`);
}

function resolveRepoPath(repoRoot, relativePath) {
    const normalizedRelativePath = String(relativePath).trim().replace(/^['"]|['"]$/g, '');
    return assertPathWithinRepo(repoRoot, path.resolve(repoRoot, normalizedRelativePath));
}

function readJsonFile(filePath) {
    return JSON.parse(fs.readFileSync(filePath, 'utf8'));
}

function readYamlFile(filePath) {
    const text = fs.readFileSync(filePath, 'utf8');
    const documents = yaml.parseAllDocuments(text);
    for (const document of documents) {
        if (document.errors.length > 0) {
            throw new Error(
                `${filePath}: ${document.errors.map((error) => error.message).join('; ')}`,
            );
        }
    }
    return documents.filter((document) => document.contents !== null).map((document) => document.toJSON());
}

function asArray(value) {
    if (value === undefined || value === null) {
        return [];
    }
    return Array.isArray(value) ? value : [value];
}

function uniqueValues(values, preserveEmpty = false) {
    const seen = new Set();
    const result = [];
    for (const value of values) {
        if ((value === undefined || value === null) && !preserveEmpty) {
            continue;
        }
        const stringValue = value === undefined || value === null ? '' : String(value).trim();
        if ((!stringValue && !preserveEmpty) || seen.has(stringValue)) {
            continue;
        }
        seen.add(stringValue);
        result.push(stringValue);
    }
    return result;
}

function addSelector(selectors, selector) {
    selectors.push({
        toolchain: 'node',
        selectorClass: selector.selectorClass,
        sourceType: selector.sourceType,
        origin: selector.origin,
        path: toPosixPath(selector.path),
        rawValue: String(selector.rawValue).trim(),
        referencedPath: selector.referencedPath ? toPosixPath(selector.referencedPath) : undefined,
    });
}

function collectPackageSelectors(repoRoot, selectors, problems) {
    const packageJsonPath = path.join(repoRoot, 'package.json');
    const packageLockPath = path.join(repoRoot, 'package-lock.json');
    let packageJsonEngine;
    let packageLockEngine;

    if (fs.existsSync(packageJsonPath)) {
        const packageJson = readJsonFile(packageJsonPath);
        packageJsonEngine = packageJson.engines && packageJson.engines.node;
        if (typeof packageJsonEngine === 'string' && packageJsonEngine.trim()) {
            addSelector(selectors, {
                selectorClass: 'support-floor',
                sourceType: 'package-json:engines.node',
                origin: 'root package engines.node',
                path: 'package.json',
                rawValue: packageJsonEngine,
            });
        }
    }

    if (fs.existsSync(packageLockPath)) {
        const packageLock = readJsonFile(packageLockPath);
        packageLockEngine =
            packageLock.packages &&
            packageLock.packages[''] &&
            packageLock.packages[''].engines &&
            packageLock.packages[''].engines.node;
        if (typeof packageLockEngine === 'string' && packageLockEngine.trim()) {
            addSelector(selectors, {
                selectorClass: 'support-floor',
                sourceType: 'package-lock:root.engines.node',
                origin: 'root package-lock mirror',
                path: 'package-lock.json',
                rawValue: packageLockEngine,
            });
        }
    }

    if (
        typeof packageJsonEngine === 'string' &&
        typeof packageLockEngine === 'string' &&
        packageJsonEngine.trim() !== packageLockEngine.trim()
    ) {
        problems.push({
            path: 'package-lock.json',
            message:
                `package-lock root engines.node (${packageLockEngine}) does not match ` +
                `package.json engines.node (${packageJsonEngine}).`,
        });
    }
}

function getWorkflowFiles(repoRoot) {
    return globSync('.github/workflows/**/*.{yml,yaml}', {
        cwd: repoRoot,
        nodir: true,
        dot: true,
        windowsPathsNoEscape: true,
    }).sort();
}

function getAzurePipelineFiles(repoRoot) {
    const patterns = [
        'azure-pipelines.yml',
        'azure-pipelines.yaml',
        '.azuredevops/pipelines/**/*.{yml,yaml}',
    ];
    const files = new Set();
    for (const pattern of patterns) {
        for (const filePath of globSync(pattern, {
            cwd: repoRoot,
            nodir: true,
            dot: true,
            windowsPathsNoEscape: true,
        })) {
            files.add(filePath);
        }
    }
    return [...files].sort();
}

function isSetupNodeStep(step) {
    return (
        step &&
        typeof step === 'object' &&
        typeof step.uses === 'string' &&
        /^actions\/setup-node@/i.test(step.uses.trim())
    );
}

function isMapping(value) {
    return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function hasGithubExpression(value) {
    return typeof value === 'string' && value.includes('${{');
}

function requireStaticMatrixValue(value, depth = 0) {
    if (depth > 32 || hasGithubExpression(value)) {
        throw new Error('Node.js matrix input cannot be resolved: dynamic or excessively nested matrix value.');
    }
    if (Array.isArray(value) || isMapping(value)) {
        for (const entry of Object.values(value)) {
            requireStaticMatrixValue(entry, depth + 1);
        }
    } else if (value !== null && !['string', 'number', 'boolean'].includes(typeof value)) {
        throw new Error('Node.js matrix input cannot be resolved: unsupported static value.');
    }
}

function matrixValuesEqual(left, right) {
    if (left === right) {
        return true;
    }
    if (left === null || right === null || typeof left !== 'object' || typeof right !== 'object' ||
        Array.isArray(left) !== Array.isArray(right)) {
        return false;
    }
    const keys = Object.keys(left);
    return keys.length === Object.keys(right).length && keys.every((key) =>
        Object.hasOwn(right, key) && matrixValuesEqual(left[key], right[key]));
}

function expandGithubMatrix(matrix) {
    if (matrix === undefined) {
        return [{}];
    }
    if (!isMapping(matrix)) {
        throw new Error('Node.js matrix input cannot be resolved: expected a static matrix mapping.');
    }
    requireStaticMatrixValue(matrix);
    const axes = Object.entries(matrix).filter(([key]) => !['include', 'exclude'].includes(key));
    for (const name of ['include', 'exclude']) {
        if (Object.hasOwn(matrix, name) &&
            (!Array.isArray(matrix[name]) || !matrix[name].every(isMapping))) {
            throw new Error(`Node.js matrix input cannot be resolved: ${name} must be a list of mappings.`);
        }
    }
    let originals = [{}];
    for (const [key, values] of axes) {
        if (!Array.isArray(values) || values.length === 0) {
            throw new Error(`Node.js matrix input cannot be resolved: axis ${key} must be a nonempty static list.`);
        }
        // This local work budget is distinct from GitHub's final 256-job limit.
        if (originals.length * values.length > 4096) {
            throw new Error('Node.js matrix input exceeds the local 4096-combination expansion budget.');
        }
        originals = originals.flatMap((entry) => values.map((value) => ({ ...entry, [key]: value })));
    }
    if (axes.length === 0 && Object.hasOwn(matrix, 'include')) {
        originals = [];
    }
    originals = originals.filter((entry) => !(matrix.exclude || []).some((excluded) =>
        Object.entries(excluded).every(([key, value]) =>
            Object.hasOwn(entry, key) && matrixValuesEqual(entry[key], value))));
    const combinations = originals.map((entry) => ({ ...entry }));
    const standalone = [];
    for (const included of matrix.include || []) {
        let matched = false;
        originals.forEach((original, index) => {
            if (Object.entries(included).every(([key, value]) =>
                !Object.hasOwn(original, key) || matrixValuesEqual(original[key], value))) {
                combinations[index] = { ...combinations[index], ...included };
                matched = true;
            }
        });
        if (!matched) {
            standalone.push(included);
        }
        if (combinations.length + standalone.length > 256) {
            throw new Error('Node.js matrix input exceeds the final 256-job matrix limit.');
        }
    }
    if (combinations.length + standalone.length > 256) {
        throw new Error('Node.js matrix input exceeds the final 256-job matrix limit.');
    }
    return [...combinations, ...standalone];
}

function resolveGithubExpression(value, combination) {
    let resolved = value;
    let origin = 'literal';
    if (hasGithubExpression(value)) {
        const match = value.match(/^\s*\$\{\{\s*matrix\.([A-Za-z_][A-Za-z0-9_-]*(?:\.[A-Za-z_][A-Za-z0-9_-]*)*)\s*\}\}\s*$/);
        if (!match) {
            throw new Error('Node.js setup input cannot be resolved: only exact static matrix property expressions are supported.');
        }
        resolved = combination;
        for (const key of match[1].split('.')) {
            resolved = isMapping(resolved) && Object.hasOwn(resolved, key) ? resolved[key] : undefined;
        }
        origin = `matrix.${match[1]}`;
    }
    if (resolved !== null && resolved !== undefined && !['string', 'number', 'boolean'].includes(typeof resolved)) {
        throw new Error('Node.js setup input cannot be resolved: the selected value must be scalar.');
    }
    return {
        rawValue: resolved === null || resolved === undefined ? '' : String(resolved),
        origin,
        missingProperty: hasGithubExpression(value) && resolved === undefined,
    };
}

function readVersionFileSelector(repoRoot, filePathValue, sourcePath, sourceType, problems) {
    let referencedAbsolutePath;
    try {
        referencedAbsolutePath = resolveRepoPath(repoRoot, filePathValue);
    } catch (error) {
        problems.push({ path: sourcePath, message: error.message });
        return [];
    }

    if (!fs.existsSync(referencedAbsolutePath)) {
        problems.push({
            path: sourcePath,
            message: `referenced Node.js version file does not exist: ${filePathValue}`,
        });
        return [];
    }

    const realRepoRoot = fs.realpathSync(repoRoot);
    const realReferencedPath = fs.realpathSync(referencedAbsolutePath);
    assertPathWithinRepo(realRepoRoot, realReferencedPath);

    const referencedPath = repoRelativePath(repoRoot, referencedAbsolutePath);
    const content = fs.readFileSync(referencedAbsolutePath, 'utf8');
    let rawValue;

    if (path.basename(referencedAbsolutePath) === '.tool-versions') {
        const nodeLine = content
            .split(/\r?\n/)
            .map((line) => line.trim())
            .find((line) => line && !line.startsWith('#') && /^nodejs\s+/i.test(line));
        if (nodeLine) {
            rawValue = nodeLine.split(/\s+/)[1];
        }
    } else {
        rawValue = content
            .split(/\r?\n/)
            .map((line) => line.trim())
            .find((line) => line && !line.startsWith('#'));
    }

    if (!rawValue) {
        problems.push({
            path: referencedPath,
            message: `referenced Node.js version file does not contain a selector.`,
        });
        return [];
    }

    return [
        {
            selectorClass: 'ci-runtime',
            sourceType,
            origin: 'version-file',
            path: sourcePath,
            rawValue,
            referencedPath,
        },
    ];
}

function readGithubNodeVersionFile(repoRoot, absolutePath, seen = new Set()) {
    // Inventory the effective setup-node value; this is not a workflow-policy checker.
    const realPath = fs.realpathSync(absolutePath);
    assertPathWithinRepo(fs.realpathSync(repoRoot), realPath);
    if (!fs.statSync(realPath).isFile()) {
        throw new Error('Node.js version source must be a regular file.');
    }
    if (seen.has(realPath) || seen.size >= 32) {
        throw new Error('Node.js version inheritance is cyclic or exceeds the 32-file inventory limit.');
    }
    seen.add(realPath);
    const content = fs.readFileSync(realPath, 'utf8');
    let manifest;
    try {
        manifest = JSON.parse(content);
    } catch (error) {
        if (!(error instanceof SyntaxError)) {
            throw error;
        }
        if (path.basename(realPath) === 'package.json') {
            throw new Error('Node.js version package.json must contain valid JSON.');
        }
    }

    let rawValue;
    if (manifest && typeof manifest === 'object') {
        rawValue = manifest.volta && manifest.volta.node;
        if (!rawValue) {
            const runtime = asArray(manifest.devEngines && manifest.devEngines.runtime)
                .find((entry) => {
                    if (entry === null || (entry.name != null && typeof entry.name !== 'string')) {
                        throw new Error('devEngines.runtime entries require a string name.');
                    }
                    return typeof entry.name === 'string' &&
                        entry.name.toLowerCase() === 'node' && entry.version;
                });
            rawValue = (runtime && runtime.version) || (manifest.engines && manifest.engines.node);
        }
        if (!rawValue && manifest.volta && manifest.volta.extends) {
            if (typeof manifest.volta.extends !== 'string') {
                throw new Error('volta.extends must be a repository-contained file path.');
            }
            const inheritedPath = path.resolve(path.dirname(absolutePath), manifest.volta.extends);
            assertPathWithinRepo(repoRoot, inheritedPath);
            return readGithubNodeVersionFile(repoRoot, inheritedPath, seen);
        }
    } else {
        // Match the inspected setup-node parser, including node/nodejs prefixes.
        const match = content.match(/^(?:node(js)?\s+)?v?(?<version>[^\s]+)$/m);
        rawValue = match ? match.groups.version : content.trim();
    }
    if (typeof rawValue !== 'string' || !rawValue.trim()) {
        throw new Error('Node.js version file does not select a string version.');
    }
    return { rawValue, referencedPath: repoRelativePath(repoRoot, absolutePath) };
}

function collectGithubStepSelectors(repoRoot, sourcePath, stepWith, matrix, selectors, problems) {
    const direct = stepWith['node-version'];
    const file = stepWith['node-version-file'];
    const addDirect = (resolved) => addSelector(selectors, {
        selectorClass: 'ci-runtime',
        sourceType: 'github-actions:setup-node node-version',
        origin: resolved.origin,
        path: sourcePath,
        rawValue: resolved.rawValue,
    });
    // A literal direct version does not consult the matrix or version file.
    if (!hasGithubExpression(direct)) {
        const resolved = resolveGithubExpression(direct, {});
        if (resolved.rawValue.trim()) {
            addDirect(resolved);
            return;
        }
    }
    const combinations = hasGithubExpression(direct) || hasGithubExpression(file)
        ? expandGithubMatrix(matrix) : [{}];
    for (const combination of combinations) {
        try {
            const resolved = resolveGithubExpression(direct, combination);
            if (resolved.rawValue.trim()) {
                addDirect(resolved);
                continue;
            }
            if (!Object.hasOwn(stepWith, 'node-version-file')) {
                if (resolved.missingProperty) {
                    throw new Error(`Node.js direct matrix input ${resolved.origin} is missing and has no version-file fallback; no checked-in runtime can be inventoried.`);
                }
                throw new Error('Node.js setup input has no nonblank node-version or version-file fallback; no checked-in runtime can be inventoried.');
            }
            const selectedFile = resolveGithubExpression(file, combination);
            if (!selectedFile.rawValue.trim()) {
                throw new Error('Node.js version-file input has a blank value; no checked-in runtime can be inventoried.');
            }
            const selected = readGithubNodeVersionFile(
                repoRoot,
                resolveRepoPath(repoRoot, selectedFile.rawValue),
            );
            addSelector(selectors, {
                selectorClass: 'ci-runtime',
                sourceType: 'github-actions:setup-node node-version-file',
                origin: 'version-file',
                path: sourcePath,
                ...selected,
            });
        } catch (error) {
            problems.push({ path: sourcePath, message: error.message });
        }
    }
}

function collectGithubWorkflowSelectors(repoRoot, selectors, problems) {
    const collected = [];
    const issues = [];
    for (const relativeWorkflowPath of getWorkflowFiles(repoRoot)) {
        const workflowPath = path.join(repoRoot, relativeWorkflowPath);
        for (const workflow of readYamlFile(workflowPath)) {
            if (!workflow || typeof workflow !== 'object' || !workflow.jobs) {
                continue;
            }
            for (const job of Object.values(workflow.jobs)) {
                if (!job || typeof job !== 'object') {
                    continue;
                }
                const matrix = job.strategy && job.strategy.matrix;
                for (const step of asArray(job.steps)) {
                    if (!isSetupNodeStep(step)) {
                        continue;
                    }
                    try {
                        collectGithubStepSelectors(repoRoot, relativeWorkflowPath, step.with || {}, matrix, collected, issues);
                    } catch (error) {
                        issues.push({ path: relativeWorkflowPath, message: error.message });
                    }
                }
            }
        }
    }
    // Different jobs can select the same runtime and produce the same diagnostic.
    for (const [records, destination] of [[collected, selectors], [issues, problems]]) {
        const seen = new Set();
        for (const record of records) {
            const key = JSON.stringify(record);
            if (!seen.has(key)) {
                destination.push(record);
                seen.add(key);
            }
        }
    }
}

function collectParameters(parameters) {
    const result = new Map();
    for (const parameter of asArray(parameters)) {
        if (!parameter || typeof parameter !== 'object' || !parameter.name) {
            continue;
        }
        result.set(String(parameter.name), {
            defaultValue: parameter.default,
            values: asArray(parameter.values),
        });
    }
    return result;
}

function mergeMaps(...maps) {
    const merged = new Map();
    for (const map of maps) {
        for (const [key, value] of map.entries()) {
            merged.set(key, value);
        }
    }
    return merged;
}

function collectVariables(variables) {
    const result = new Map();
    if (!variables) {
        return result;
    }

    if (Array.isArray(variables)) {
        for (const variable of variables) {
            if (variable && typeof variable === 'object' && variable.name && variable.value !== undefined) {
                result.set(String(variable.name), variable.value);
            }
        }
        return result;
    }

    if (typeof variables === 'object') {
        for (const [key, value] of Object.entries(variables)) {
            if (value && typeof value === 'object' && Object.prototype.hasOwnProperty.call(value, 'value')) {
                result.set(key, value.value);
            } else {
                result.set(key, value);
            }
        }
    }
    return result;
}

function resolveAzureSelectorValue(value, context, sourcePath, problems, active = []) {
    if (context.inventoryBudget && ++context.inventoryBudget.nodes > AZURE_INVENTORY_LIMITS.nodes) {
        throw new Error('Azure inventory exceeds the 200000-node work limit.');
    }
    const fail = (message) => { problems.push({ path: sourcePath, message }); return []; };
    if (Array.isArray(value)) {
        return value.flatMap((entry) => resolveAzureSelectorValue(entry, context, sourcePath, problems, active));
    }
    if (value === null || value === undefined || typeof value === 'object') {
        return fail('Azure Node selector or template argument must resolve to a checked-in scalar.');
    }
    if (typeof value !== 'string') return [{ rawValue: value, origin: 'literal' }];
    const parameterMatch = value.match(/^\s*\$\{\{\s*parameters\.([A-Za-z0-9_.-]+)\s*\}\}\s*$/);
    const variableMatch = value.match(/^\s*\$\(([A-Za-z0-9_.-]+)\)\s*$/) ||
        value.match(/^\s*\$\{\{\s*variables\.([A-Za-z0-9_.-]+)\s*\}\}\s*$/);
    let values;
    let origin;
    if (parameterMatch) {
        const parameter = context.parameters.get(parameterMatch[1]);
        // An absent default is not an empty alternative; explicit values remain
        // intact until the selected task validates them.
        values = parameter ? [
            ...(parameter.defaultValue === undefined ? [] : [parameter.defaultValue]),
            ...parameter.values,
        ] : [];
        origin = `parameters.${parameterMatch[1]} ${parameter && parameter.values.length ? 'default-or-values' : 'default'}`;
    } else if (variableMatch) {
        const name = variableMatch[1];
        if (context.strategy && Object.hasOwn(context.strategy, 'matrix')) {
            const matrix = context.strategy.matrix;
            if (!isMapping(matrix) || Object.keys(matrix).length === 0) {
                return fail('Azure selected variable matrix must be a nonempty checked-in mapping.');
            }
            const resolved = [];
            for (const [leg, entry] of Object.entries(matrix)) {
                if (!isMapping(entry) || leg.includes('$')) {
                    fail('Azure selected matrix leg must be a checked-in variable mapping.');
                    continue;
                }
                // Charge map copying as work as well as recursive resolution.
                if (context.inventoryBudget) {
                    context.inventoryBudget.nodes += context.variables.size + Object.keys(entry).length;
                    if (context.inventoryBudget.nodes > AZURE_INVENTORY_LIMITS.nodes) {
                        throw new Error('Azure inventory exceeds the 200000-node work limit.');
                    }
                }
                resolved.push(...resolveAzureSelectorValue(value, {
                    ...context,
                    variables: mergeMaps(context.variables, new Map(Object.entries(entry))),
                    strategy: undefined,
                    matrixEntry: entry,
                }, sourcePath, problems, active));
            }
            return resolved;
        }
        const variable = context.variables.get(name);
        if (context.matrixEntry && Object.hasOwn(context.matrixEntry, name) &&
            variable !== null && typeof variable === 'object') {
            return fail('Azure selected matrix variable must be a checked-in scalar.');
        }
        values = asArray(variable);
        origin = `variable-or-matrix.${name}`;
    } else {
        if (/\$\{|\$\(|\$\[/.test(value)) return fail(`Azure selector expression cannot be verified from checked-in YAML: ${value}`);
        return [{ rawValue: value, origin: 'literal' }];
    }
    if (values.length === 0) return fail(`Azure selector ${value.trim()} cannot be verified from checked-in YAML; no checked-in value is available.`);
    const identity = parameterMatch ? `parameter:${parameterMatch[1]}` : `variable:${variableMatch[1]}`;
    if (active.includes(identity) || active.length >= 100) return fail('Azure selector references are cyclic or exceed the 100-level nesting limit.');
    return values.flatMap((rawValue) => {
        if (parameterMatch && Array.isArray(rawValue)) {
            return fail('Azure parameter selector alternative must resolve to a checked-in scalar.');
        }
        return resolveAzureSelectorValue(rawValue, context, sourcePath, problems, [...active, identity]);
    })
        .map((resolved) => ({ ...resolved, origin }));
}

function resolveAzureNodeSelector(value, context, sourcePath, problems) {
    return resolveAzureSelectorValue(value, context, sourcePath, problems).filter((resolved) => {
        if (typeof resolved.rawValue === 'string' && !resolved.rawValue.trim()) {
            problems.push({
                path: sourcePath,
                message: 'Azure Node selector must be a nonblank checked-in scalar.',
            });
            return false;
        }
        return true;
    });
}

function isAzureNodeTask(step) {
    if (!step || typeof step !== 'object' || typeof step.task !== 'string') {
        return false;
    }
    return /^(UseNode@1|NodeTool@0)$/i.test(step.task.trim());
}

function collectAzureStepSelectors(repoRoot, relativePipelinePath, steps, context, selectors, problems) {
    for (const step of asArray(steps)) {
        if (!isAzureNodeTask(step)) {
            continue;
        }

        const taskName = step.task.trim();
        const inputs = step.inputs || {};
        const selectedInput = /^UseNode@1$/i.test(taskName) ? 'version'
            : /^fromFile$/i.test(String(inputs.versionSource || '').trim())
                ? 'versionFilePath' : 'versionSpec';
        if (!Object.hasOwn(inputs, selectedInput)) {
            problems.push({
                path: relativePipelinePath,
                message: `Azure Pipelines ${taskName} has no checked-in ${selectedInput} input; no runtime can be inventoried.`,
            });
            continue;
        }
        if (/^UseNode@1$/i.test(taskName) && Object.prototype.hasOwnProperty.call(inputs, 'version')) {
            for (const resolved of resolveAzureNodeSelector(inputs.version, context, relativePipelinePath, problems)) {
                addSelector(selectors, {
                    selectorClass: 'ci-runtime',
                    sourceType: 'azure-pipelines:UseNode@1 version',
                    origin: resolved.origin,
                    path: relativePipelinePath,
                    rawValue: resolved.rawValue,
                });
            }
        }

        if (/^NodeTool@0$/i.test(taskName)) {
            const versionSource = String(inputs.versionSource || '').trim();
            if (
                /^fromFile$/i.test(versionSource) &&
                Object.prototype.hasOwnProperty.call(inputs, 'versionFilePath')
            ) {
                for (const resolved of resolveAzureNodeSelector(inputs.versionFilePath, context, relativePipelinePath, problems)) {
                    for (const selector of readVersionFileSelector(
                        repoRoot,
                        resolved.rawValue,
                        relativePipelinePath,
                        'azure-pipelines:NodeTool@0 versionFilePath',
                        problems,
                    )) {
                        addSelector(selectors, selector);
                    }
                }
            } else if (Object.prototype.hasOwnProperty.call(inputs, 'versionSpec')) {
                for (const resolved of resolveAzureNodeSelector(inputs.versionSpec, context, relativePipelinePath, problems)) {
                    addSelector(selectors, {
                        selectorClass: 'ci-runtime',
                        sourceType: 'azure-pipelines:NodeTool@0 versionSpec',
                        origin: resolved.origin,
                        path: relativePipelinePath,
                        rawValue: resolved.rawValue,
                    });
                }
            }
        }
    }
}

// These are local inventory limits, not Azure's complete compiler contract.
const AZURE_INVENTORY_LIMITS = Object.freeze({
    files: 100, depth: 100, invocations: 4096,
    fileBytes: 2 * 1024 * 1024, totalBytes: 20 * 1024 * 1024, nodes: 200000,
});

function azureDocumentSize(documents) {
    let nodes = 0;
    const active = new Set();
    const stack = documents.map((value) => ({ value, exit: false }));
    while (stack.length > 0) {
        const { value, exit } = stack.pop();
        if (exit) {
            active.delete(value);
            continue;
        }
        nodes += 1;
        if (nodes > AZURE_INVENTORY_LIMITS.nodes) {
            throw new Error('Azure inventory exceeds the 200000-node work limit.');
        }
        if (value && typeof value === 'object') {
            if (active.has(value)) throw new Error('Azure YAML contains a cyclic alias.');
            active.add(value);
            stack.push({ value, exit: true });
            for (const child of Object.values(value)) stack.push({ value: child, exit: false });
        }
    }
    return nodes;
}

function collectAzurePipelineSelectors(repoRoot, selectors, problems, explicitPaths = []) {
    if (!Array.isArray(explicitPaths)) throw new Error('Azure pipeline paths must be an array.');
    const realRoot = fs.realpathSync(repoRoot);
    const cache = new Map();
    const referenced = new Set();
    const addProblem = (source, error) => problems.push({ path: source, message: error.message });
    const newBudget = () => ({ files: new Set(), bytes: 0, nodes: 0, invocations: 0 });
    const pathName = (absolute) => repoRelativePath(repoRoot, absolute);

    function localPath(value, including, entrypoint = false) {
        if (typeof value !== 'string' || !value.trim() || value !== value.trim() ||
            /[\x00-\x1f*?{}$\\]/.test(value)) {
            throw new Error('Azure template or pipeline path must be a nonblank literal local YAML path.');
        }
        let filename = value;
        if (filename.includes('@')) {
            if (entrypoint || !filename.endsWith('@self') || filename.slice(0, -5).includes('@')) {
                throw new Error(`Azure external template cannot be verified locally: ${value}`);
            }
            filename = filename.slice(0, -5);
        }
        if (entrypoint && (filename.startsWith('/') || path.isAbsolute(filename))) {
            throw new Error('An explicit Azure pipeline path must be repository-relative.');
        }
        if (!/\.ya?ml$/i.test(filename) || /^[A-Za-z]:/.test(filename)) {
            throw new Error('Azure template or pipeline source must be a local .yml or .yaml file.');
        }
        const absolute = filename.startsWith('/')
            ? path.resolve(repoRoot, filename.slice(1))
            : path.resolve(including ? path.dirname(including) : repoRoot, filename);
        assertPathWithinRepo(repoRoot, absolute);
        const real = fs.realpathSync(absolute);
        assertPathWithinRepo(realRoot, real);
        if (!fs.statSync(real).isFile()) throw new Error('Azure YAML source must be a regular file.');
        return { absolute, real };
    }

    function load(file, budget, active) {
        budget.invocations += 1;
        if (budget.invocations > AZURE_INVENTORY_LIMITS.invocations) {
            throw new Error('Azure inventory exceeds the 4096-invocation work limit.');
        }
        if (active.includes(file.real)) throw new Error('Azure template references form a cycle.');
        if (active.length >= AZURE_INVENTORY_LIMITS.depth) {
            throw new Error('Azure inventory exceeds the 100-level nesting limit.');
        }
        if (!cache.has(file.real)) {
            if (fs.statSync(file.real).size > AZURE_INVENTORY_LIMITS.fileBytes) {
                throw new Error('Azure YAML source exceeds the 2 MiB file limit.');
            }
            const bytes = fs.readFileSync(file.real);
            if (bytes.length > AZURE_INVENTORY_LIMITS.fileBytes) {
                throw new Error('Azure YAML source exceeds the 2 MiB file limit.');
            }
            const documents = yaml.parseAllDocuments(new TextDecoder('utf-8', { fatal: true }).decode(bytes));
            const invalid = documents.flatMap((document) => document.errors);
            if (invalid.length > 0) throw new Error(`Invalid Azure YAML: ${invalid.map((error) => error.message).join('; ')}`);
            const values = documents.filter((document) => document.contents !== null).map((document) => document.toJSON());
            if (values.length === 0 || !values.every(isMapping)) {
                throw new Error('Azure YAML source must contain a pipeline or template mapping.');
            }
            cache.set(file.real, { documents: values, bytes: bytes.length, nodes: azureDocumentSize(values) });
        }
        const loaded = cache.get(file.real);
        if (!budget.files.has(file.real)) {
            budget.files.add(file.real);
            budget.bytes += loaded.bytes;
        }
        if (budget.files.size > AZURE_INVENTORY_LIMITS.files) {
            throw new Error('Azure inventory exceeds the 100-file limit.');
        }
        if (budget.bytes > AZURE_INVENTORY_LIMITS.totalBytes) {
            throw new Error('Azure inventory exceeds the 20 MiB cumulative input limit.');
        }
        budget.nodes += loaded.nodes;
        if (budget.nodes > AZURE_INVENTORY_LIMITS.nodes) {
            throw new Error('Azure inventory exceeds the 200000-node work limit.');
        }
        return loaded.documents;
    }

    function visit(file, role, overrides, inherited, budget, active, discovery) {
        const source = pathName(file.absolute);
        const documents = load(file, budget, active);
        const ancestry = [...active, file.real];
        let returnedVariables = inherited.variables;
        for (const document of documents) {
            const parameters = collectParameters(document.parameters);
            if (overrides !== null && !discovery) {
                // Allowed parameter values are not invocations of a called template.
                for (const parameter of parameters.values()) parameter.values = [];
                for (const [name, value] of Object.entries(overrides)) {
                    if (!parameters.has(name)) {
                        addProblem(source, new Error(`Azure template argument "${name}" has no declaration.`));
                        continue;
                    }
                    if (value === null || typeof value === 'object') {
                        addProblem(source, new Error(`Azure template argument "${name}" requires a supported scalar binding.`));
                        parameters.set(name, { defaultValue: undefined, values: [] });
                        continue;
                    }
                    const resolved = resolveAzureSelectorValue(value, inherited, source, problems)
                        .map((entry) => entry.rawValue);
                    parameters.set(name, { defaultValue: resolved[0], values: resolved.slice(1) });
                }
            }
            const context = { parameters, variables: new Map(inherited.variables), strategy: inherited.strategy, inventoryBudget: budget };

            function template(reference, templateRole, caller) {
                try {
                    if (!isMapping(reference) || !Object.hasOwn(reference, 'template')) {
                        throw new Error('Azure template reference must contain a literal template path.');
                    }
                    const selected = localPath(reference.template, file.absolute);
                    if (discovery) referenced.add(selected.real);
                    const argumentsMap = reference.parameters === undefined ? {} : reference.parameters;
                    if (!isMapping(argumentsMap)) throw new Error('Azure template parameters must be a mapping.');
                    return visit(selected, templateRole, argumentsMap, caller, budget, ancestry, discovery);
                } catch (error) {
                    addProblem(source, error);
                    return new Map(caller.variables);
                }
            }

            function variables(raw, caller) {
                const result = new Map(caller.variables);
                if (raw === undefined) return result;
                const declaredValue = (value) => {
                    if (!discovery && typeof value === 'string' &&
                        /^\s*\$\{\{\s*(?:parameters|variables)\.[A-Za-z0-9_.-]+\s*\}\}\s*$/.test(value)) {
                        // Compile-time expressions belong to this declaration, not a callee's parameters.
                        return resolveAzureSelectorValue(value, { ...caller, variables: result }, source, [])
                            .map((entry) => entry.rawValue);
                    }
                    return value;
                };
                if (!Array.isArray(raw)) {
                    if (!isMapping(raw)) {
                        addProblem(source, new Error('Azure variables require a static mapping or list.'));
                        return result;
                    }
                    for (const [name, value] of collectVariables(raw)) {
                        if (name.includes('${{')) addProblem(source, new Error('Azure conditional variable assembly cannot be verified locally.'));
                        else result.set(name, declaredValue(value));
                    }
                    return result;
                }
                for (const entry of raw) {
                    if (!isMapping(entry)) {
                        addProblem(source, new Error('Azure variable entry cannot be verified locally.'));
                    } else if (Object.hasOwn(entry, 'template')) {
                        for (const [name, value] of template(entry, 'variables', { ...caller, variables: result })) result.set(name, value);
                    } else if (Object.keys(entry).some((key) => key.includes('${{'))) {
                        addProblem(source, new Error('Azure conditional variable assembly cannot be verified locally.'));
                    } else if (entry.name && Object.hasOwn(entry, 'value')) {
                        result.set(String(entry.name), declaredValue(entry.value));
                    } else if (!Object.hasOwn(entry, 'group')) {
                        addProblem(source, new Error('Azure variable entry cannot be verified locally.'));
                    }
                    // External variable groups never supply checked-in values.
                }
                return result;
            }

            function scope(value, caller) {
                const local = { ...caller, variables: variables(value.variables, caller), strategy: value.strategy || caller.strategy };
                if (Object.hasOwn(value, 'extends')) template(value.extends, 'pipeline', local);
                for (const childRole of ['steps', 'jobs', 'stages']) {
                    if (Object.hasOwn(value, childRole)) sequence(value[childRole], childRole, local);
                }
                return local.variables;
            }

            function sequence(raw, childRole, caller) {
                if (!Array.isArray(raw)) {
                    addProblem(source, new Error(`Azure ${childRole} assembly requires a static list; expressions cannot be verified locally.`));
                    return;
                }
                for (const entry of raw) {
                    if (!isMapping(entry) || Object.keys(entry).some((key) => key.includes('${{'))) {
                        addProblem(source, new Error(`Azure conditional or structural ${childRole} assembly cannot be verified locally.`));
                    } else if (Object.hasOwn(entry, 'template')) {
                        template(entry, childRole, caller);
                    } else if (childRole === 'steps') {
                        if (!discovery) collectAzureStepSelectors(repoRoot, source, [entry], caller, selectors, problems);
                    } else {
                        if (Object.hasOwn(entry, 'deployment')) {
                            addProblem(source, new Error('Azure deployment strategy assembly cannot be verified locally.'));
                        }
                        scope(entry, caller);
                    }
                }
            }

            if (role !== 'pipeline' && !Object.hasOwn(document, role)) {
                addProblem(source, new Error(`Azure ${role} template has no ${role} declaration.`));
                continue;
            }
            if (role === 'variables') {
                returnedVariables = variables(document.variables, context);
            } else {
                returnedVariables = scope(document, context);
            }
        }
        return returnedVariables;
    }

    const candidates = new Map();
    const forcedRoots = new Set();
    for (const value of [...getAzurePipelineFiles(repoRoot).map(toPosixPath), ...explicitPaths]) {
        try {
            const file = localPath(value, null, true);
            candidates.set(file.real, file);
            if (explicitPaths.includes(value) || /^azure-pipelines\.ya?ml$/.test(value)) forcedRoots.add(file.real);
        } catch (error) {
            addProblem(String(value), error);
        }
    }
    const emptyContext = () => ({ parameters: new Map(), variables: new Map(), strategy: undefined });
    // Discover roles first so an unused template default is not a separate runtime.
    // Discovery errors, including rootless cycles, cannot disappear with role filtering.
    for (const file of candidates.values()) {
        try { visit(file, 'pipeline', null, emptyContext(), newBudget(), [], true); }
        catch (error) { addProblem(pathName(file.absolute), error); }
    }
    for (const file of candidates.values()) {
        if (!forcedRoots.has(file.real) && referenced.has(file.real)) continue;
        try { visit(file, 'pipeline', null, emptyContext(), newBudget(), [], false); }
        catch (error) { addProblem(pathName(file.absolute), error); }
    }
    for (const records of [selectors, problems]) {
        const unique = new Map(records.map((record) => [JSON.stringify(record), record]));
        records.splice(0, records.length, ...unique.values());
    }
}

function collectNodeSelectors(repoRoot = process.cwd(), options = {}) {
    const resolvedRepoRoot = path.resolve(repoRoot);
    const selectors = [];
    const problems = [];

    collectPackageSelectors(resolvedRepoRoot, selectors, problems);
    collectGithubWorkflowSelectors(resolvedRepoRoot, selectors, problems);
    collectAzurePipelineSelectors(resolvedRepoRoot, selectors, problems, options.azurePipelinePaths || []);

    return { selectors, problems };
}

function selectorReleaseLine(selector) {
    const rawValue = selector.rawValue.replace(/^v(?=\d)/i, '').trim();
    const minimumVersion = semver.minVersion(rawValue);
    if (!minimumVersion) {
        return null;
    }
    return minimumVersion.major;
}

function parseUtcDateOnly(dateText) {
    const match = String(dateText).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!match) {
        throw new Error(`expected date in YYYY-MM-DD form: ${dateText}`);
    }
    return new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
}

function normalizeNodeSchedule(scheduleJson) {
    const releases = new Map();
    for (const [key, value] of Object.entries(scheduleJson)) {
        const match = key.match(/^v?(\d+)(?:\.\d+)?$/);
        if (!match || !value || typeof value !== 'object' || !value.end) {
            continue;
        }
        releases.set(Number(match[1]), {
            releaseLine: Number(match[1]),
            eolDate: String(value.end),
        });
    }
    return releases;
}

function evaluateSelectors(selectors, scheduleJson, options = {}) {
    const asOfDate = parseUtcDateOnly(
        options.asOfDate || new Date().toISOString().slice(0, 10),
    );
    const warningWindowDays = Number(options.warningWindowDays ?? DEFAULT_WARNING_WINDOW_DAYS);
    const schedule = normalizeNodeSchedule(scheduleJson);
    const findings = [];
    const problems = [];

    for (const selector of selectors) {
        if (!['ci-runtime', 'support-floor'].includes(selector.selectorClass)) {
            continue;
        }
        const releaseLine = selectorReleaseLine(selector);
        if (releaseLine === null) {
            problems.push({
                path: selector.path,
                message: `unable to parse Node.js selector "${selector.rawValue}" from ${selector.sourceType}.`,
            });
            continue;
        }

        const scheduleEntry = schedule.get(releaseLine);
        if (!scheduleEntry) {
            problems.push({
                path: selector.path,
                message:
                    `Node.js ${releaseLine} from selector "${selector.rawValue}" is not present ` +
                    'in the Node.js release schedule.',
            });
            continue;
        }

        const eolDate = parseUtcDateOnly(scheduleEntry.eolDate);
        const daysUntilEol = Math.floor((eolDate.getTime() - asOfDate.getTime()) / MILLIS_PER_DAY);
        let status = 'supported';
        if (daysUntilEol <= 0) {
            // The schedule's end date is the EOL date itself, so "at EOL"
            // (daysUntilEol === 0) gets the same stronger signal as "past EOL".
            status = 'eol';
        } else if (daysUntilEol <= warningWindowDays) {
            status = 'near-eol';
        }

        findings.push({
            ...selector,
            releaseLine,
            eolDate: scheduleEntry.eolDate,
            daysUntilEol,
            status,
        });
    }

    return { findings, problems };
}

async function loadSchedule(options) {
    if (options.scheduleFile) {
        return readJsonFile(path.resolve(options.scheduleFile));
    }

    const scheduleUrl = options.scheduleUrl || DEFAULT_NODE_SCHEDULE_URL;
    const timeoutMs = Number(options.fetchTimeoutMs ?? DEFAULT_FETCH_TIMEOUT_MS);
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    try {
        const response = await fetch(scheduleUrl, { signal: controller.signal });
        if (!response.ok) {
            throw new Error(`failed to fetch Node.js release schedule from ${scheduleUrl}: ${response.status}`);
        }
        return await response.json();
    } catch (error) {
        if (error && error.name === 'AbortError') {
            throw new Error(
                `timed out after ${timeoutMs}ms fetching Node.js release schedule from ${scheduleUrl}.`,
            );
        }
        throw error;
    } finally {
        clearTimeout(timer);
    }
}

function parseArgs(argv) {
    const options = {
        repoRoot: process.cwd(),
        scheduleUrl: DEFAULT_NODE_SCHEDULE_URL,
        warningWindowDays: Number(
            process.env.TOOLCHAIN_EOL_WARNING_DAYS || DEFAULT_WARNING_WINDOW_DAYS,
        ),
        json: false,
        azurePipelinePaths: [],
    };

    for (let index = 0; index < argv.length; index++) {
        const arg = argv[index];
        if (arg === '--repo-root') {
            options.repoRoot = argv[++index];
        } else if (arg === '--azure-pipeline') {
            const selected = argv[++index];
            if (!selected || !selected.trim() || selected.startsWith('--')) {
                throw new Error('--azure-pipeline requires a repository-relative YAML path.');
            }
            options.azurePipelinePaths.push(selected);
        } else if (arg === '--schedule-file') {
            options.scheduleFile = argv[++index];
        } else if (arg === '--schedule-url') {
            options.scheduleUrl = argv[++index];
        } else if (arg === '--warning-window-days') {
            options.warningWindowDays = Number(argv[++index]);
        } else if (arg === '--as-of') {
            options.asOfDate = argv[++index];
        } else if (arg === '--json') {
            options.json = true;
        } else if (arg === '--help' || arg === '-h') {
            options.help = true;
        } else {
            throw new Error(`unknown argument: ${arg}`);
        }
    }

    if (!Number.isInteger(options.warningWindowDays) || options.warningWindowDays < 0) {
        throw new Error('--warning-window-days must be a non-negative integer.');
    }

    return options;
}

function printHelp() {
    console.log(`Usage: node .github/scripts/check-toolchain-eol.js [options]

Options:
  --repo-root PATH             Repository root to scan. Defaults to cwd.
  --azure-pipeline PATH        Add a custom Azure entrypoint (repeatable).
  --schedule-file PATH         Read Node.js release schedule JSON from a fixture file.
  --schedule-url URL           Fetch Node.js release schedule JSON from URL.
  --warning-window-days DAYS   Warn/fail when EOL is within DAYS. Default: 180.
  --as-of YYYY-MM-DD           Evaluate against this UTC date. Defaults to today.
  --json                       Emit machine-readable JSON.
`);
}

function githubAnnotationLevel(finding) {
    return finding.status === 'supported' ? 'notice' : 'error';
}

function escapeWorkflowData(value) {
    // Percent MUST be escaped first; otherwise a literal "%0A" in the text would
    // be decoded as a newline by the runner command processor (and our own
    // inserted %0D/%0A escapes would be double-encoded).
    return String(value)
        .replace(/%/g, '%25')
        .replace(/\r/g, '%0D')
        .replace(/\n/g, '%0A');
}

function escapeWorkflowProperty(value) {
    return escapeWorkflowData(value)
        .replace(/:/g, '%3A')
        .replace(/,/g, '%2C');
}

function printGithubAnnotation(level, pathName, message) {
    console.log(`::${level} file=${escapeWorkflowProperty(pathName)}::${escapeWorkflowData(message)}`);
}

function renderTextReport(result, options) {
    const monitoredFindings = result.findings.filter((finding) =>
        ['ci-runtime', 'support-floor'].includes(finding.selectorClass),
    );
    const failingFindings = monitoredFindings.filter((finding) => finding.status !== 'supported');

    console.log('Toolchain EOL monitor');
    console.log(`Data source: ${options.scheduleFile || options.scheduleUrl || DEFAULT_NODE_SCHEDULE_URL}`);
    console.log(`Warning window: ${options.warningWindowDays} day(s)`);
    console.log('');

    if (monitoredFindings.length === 0) {
        console.log('No monitored Node.js selectors were discovered.');
    } else {
        console.log('Monitored Node.js selectors:');
        for (const finding of monitoredFindings) {
            const detail = finding.referencedPath ? ` via ${finding.referencedPath}` : '';
            console.log(
                `- ${finding.selectorClass} Node.js ${finding.releaseLine} ` +
                    `(${finding.rawValue}) in ${finding.path}${detail}: ` +
                    `${finding.status}; EOL ${finding.eolDate}; ` +
                    `${finding.daysUntilEol} day(s) remaining`,
            );
        }
    }

    if (result.problems.length > 0) {
        console.log('');
        console.log('Inventory problems:');
        for (const problem of result.problems) {
            console.log(`- ${problem.path}: ${problem.message}`);
            printGithubAnnotation('error', problem.path, problem.message);
        }
    }

    for (const finding of failingFindings) {
        const message =
            finding.status === 'eol'
                ? `Node.js ${finding.releaseLine} is past EOL (${finding.eolDate}).`
                : `Node.js ${finding.releaseLine} reaches EOL on ${finding.eolDate}, ` +
                  `within the ${options.warningWindowDays}-day warning window.`;
        printGithubAnnotation(githubAnnotationLevel(finding), finding.path, message);
    }

    if (result.problems.length === 0 && failingFindings.length === 0) {
        console.log('');
        console.log('All monitored Node.js selectors are outside the configured warning window.');
    }
}

async function runCli(argv = process.argv.slice(2)) {
    const options = parseArgs(argv);
    if (options.help) {
        printHelp();
        return 0;
    }

    const inventory = collectNodeSelectors(options.repoRoot, options);
    const schedule = await loadSchedule(options);
    const evaluation = evaluateSelectors(inventory.selectors, schedule, options);
    const result = {
        selectors: inventory.selectors,
        findings: evaluation.findings,
        problems: [...inventory.problems, ...evaluation.problems],
    };

    if (options.json) {
        console.log(JSON.stringify(result, null, 2));
    } else {
        renderTextReport(result, options);
    }

    const hasFailingFinding = result.findings.some((finding) => finding.status !== 'supported');
    return result.problems.length > 0 || hasFailingFinding ? 1 : 0;
}

if (require.main === module) {
    runCli()
        .then((exitCode) => {
            process.exit(exitCode);
        })
        .catch((error) => {
            console.error(`Error: ${error.message}`);
            process.exit(1);
        });
}

module.exports = {
    DEFAULT_NODE_SCHEDULE_URL,
    DEFAULT_WARNING_WINDOW_DAYS,
    collectNodeSelectors,
    escapeWorkflowData,
    evaluateSelectors,
    loadSchedule,
    normalizeNodeSchedule,
    parseArgs,
    runCli,
    selectorReleaseLine,
};
