const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawnSync } = require('node:child_process');
const test = require('node:test');

const linterPath = path.resolve(__dirname, 'lint-nested-markdown.js');
const repoRoot = path.resolve(__dirname, '../..');
const linter = require(linterPath);

function nestedDocument(depth, body = 'text\n') {
    for (let level = 0; level < depth; level++) {
        const fence = '`'.repeat(level + 3);
        body = `${fence}markdown\n${body}${fence}\n`;
    }
    return body;
}

test('nested Markdown budgets accept boundaries and reject excessive work', () => {
    assert.deepEqual(linter.extractMarkdownFencesRecursive('x'.repeat(1024 * 1024), 'bytes.md'), []);
    assert.throws(() => linter.extractMarkdownFencesRecursive('x'.repeat(1024 * 1024 + 1), 'bytes.md'), /source bytes limit/);
    assert.equal(linter.extractMarkdownFencesRecursive(nestedDocument(64), 'depth.md').length, 64);
    assert.throws(() => linter.extractMarkdownFencesRecursive(nestedDocument(65), 'depth.md'), /nesting depth limit/);
    const block = '```md\nx\n```\n';
    assert.equal(linter.extractMarkdownFencesRecursive(block.repeat(1024), 'blocks.md').length, 1024);
    assert.throws(() => linter.extractMarkdownFencesRecursive(block.repeat(1025), 'blocks.md'), /block count limit/);
    assert.equal(linter.extractMarkdownFencesRecursive(block, 'total.md', 0, 0, '', { blocks: 0, bytes: 8 * 1024 * 1024 - 2 }).length, 1);
    assert.throws(() => linter.extractMarkdownFencesRecursive(block, 'total.md', 0, 0, '', { blocks: 0, bytes: 8 * 1024 * 1024 - 1 }), /extracted bytes limit/);
    assert.throws(() => linter.extractMarkdownFencesRecursive(nestedDocument(16, 'x'.repeat(600000) + '\n'), 'overlap.md'), /extracted bytes limit/);
});

test('oversized file is rejected before reading and CLI reports budget failure', (t) => {
    const isolated = makeIsolatedCliRepository(t, 'nested-markdown-budget-');
    const input = writeFile(isolated.root, 'large.md', 'x'.repeat(1024 * 1024 + 1));
    let reads = 0;
    assert.throws(() => linter.extractMarkdownFences(input, isolated.root, {
        ...fs,
        readFileSync: () => { reads++; throw new Error('must not read'); }
    }), /source bytes limit/);
    assert.equal(reads, 0);
    const result = spawnSync(process.execPath, [isolated.executable, input], {
        cwd: isolated.root, encoding: 'utf8', timeout: 30000,
        env: { ...process.env, NODE_PATH: path.join(repoRoot, 'node_modules') }
    });
    assert.ifError(result.error);
    assert.equal(result.status, 1);
    assert.match(result.stderr, /Markdown budget exceeded.*source bytes limit/);
});

test('removing each extraction budget is caught by independent negative inputs', (t) => {
    const cases = [
        ['Buffer.byteLength(content, \'utf8\') > MAX_MARKDOWN_BYTES', "'x'.repeat(1048577)", 'source bytes'],
        ['depth >= MAX_NESTING_DEPTH', 'nestedDocument(65)', 'nesting depth'],
        ['budget.blocks > MAX_EXTRACTED_BLOCKS', "'```md\\nx\\n```\\n'.repeat(1025)", 'block count'],
        ['budget.bytes > MAX_TOTAL_EXTRACTED_BYTES', "nestedDocument(16, 'x'.repeat(600000) + '\\n')", 'extracted bytes']
    ];
    for (const [predicate, input, label] of cases) {
        const mutant = createMutant(t, predicate, 'false', label.replaceAll(' ', '-'));
        const script = [
            "const assert = require('node:assert/strict');",
            'const subject = require(process.argv[1]);',
            nestedDocument.toString(),
            `assert.throws(() => subject.extractMarkdownFencesRecursive(${input}, 'fixture.md'), /${label} limit/);`
        ].join('\n');
        assertMutantDetected(runMutant(mutant, script), label);
    }
});

function makeTempDir(t, prefix = 'nested-markdown-') {
    const directory = fs.mkdtempSync(path.join(os.tmpdir(), prefix));
    t.after(() => fs.rmSync(directory, { force: true, recursive: true }));
    return directory;
}

function makeIsolatedCliRepository(t, prefix) {
    const root = makeTempDir(t, prefix);
    const executable = path.join(root, '.github/scripts/lint-nested-markdown.js');
    fs.mkdirSync(path.dirname(executable), { recursive: true });
    fs.copyFileSync(linterPath, executable);
    return { executable, root };
}

function isWithin(parent, candidate) {
    const relative = path.relative(parent, candidate);
    return relative === '' || (!relative.startsWith(`..${path.sep}`) && relative !== '..' && !path.isAbsolute(relative));
}

test('CLI fixture placement guard identifies product-repository descendants', () => {
    assert.equal(isWithin(repoRoot, path.join(repoRoot, '.nested-markdown-would-race')), true);
});

function writeFile(root, relativePath, content = '') {
    const filePath = path.join(root, ...relativePath.split('/'));
    fs.mkdirSync(path.dirname(filePath), { recursive: true });
    fs.writeFileSync(filePath, content);
    return filePath;
}

function relativePaths(root, files) {
    return files.map((file) => path.relative(root, file).split(path.sep).join('/')).sort();
}

function tryDirectoryLink(target, linkPath) {
    try {
        fs.symlinkSync(target, linkPath, process.platform === 'win32' ? 'junction' : 'dir');
        return null;
    } catch (error) {
        if (error.code === 'EPERM' || error.code === 'EACCES' || error.code === 'ENOTSUP') {
            return error;
        }
        throw error;
    }
}

function createMutant(t, pattern, replacement, label) {
    const source = fs.readFileSync(linterPath, 'utf8');
    const changed = source.replace(pattern, replacement);
    assert.notEqual(changed, source, `Mutation did not match for ${label}.`);
    const directory = makeTempDir(t, `nested-markdown-mutant-${label}-`);
    return writeFile(directory, 'lint-nested-markdown.js', changed);
}

function runMutant(mutantPath, script, args = []) {
    return spawnSync(process.execPath, ['-e', script, mutantPath, ...args], {
        encoding: 'utf8',
        env: { ...process.env, NODE_PATH: path.join(repoRoot, 'node_modules') },
        timeout: 30000
    });
}

function assertMutantDetected(result, label) {
    assert.ifError(result.error);
    assert.equal(result.status, 1, `${label} was not detected.\nstdout:\n${result.stdout}\nstderr:\n${result.stderr}`);
    assert.match(result.stderr, /AssertionError|ERR_ASSERTION/, label);
}

test('runtime dependency failures retain the cause and actionable setup guidance', () => {
    const cause = Object.assign(new Error('synthetic missing package'), { code: 'MODULE_NOT_FOUND' });
    assert.throws(
        () => linter.loadRuntimeDependencies(() => { throw cause; }, repoRoot),
        (error) => {
            assert.equal(error.cause, cause);
            assert.match(error.message, /npm ci --ignore-scripts/);
            assert.match(error.message, /synthetic missing package/);
            return true;
        }
    );
});

test('runtime dependency loader does not relabel unrelated module errors', () => {
    const cause = Object.assign(new Error('synthetic package initialization defect'), { code: 'EINIT' });
    assert.throws(
        () => linter.loadRuntimeDependencies(() => { throw cause; }, repoRoot),
        (error) => error === cause
    );
});

test('direct execution reports dependency setup guidance when packages are unavailable', (t) => {
    const directory = makeTempDir(t, 'nested-markdown-no-dependencies-');
    const copy = path.join(directory, 'lint-nested-markdown.js');
    fs.copyFileSync(linterPath, copy);
    const result = spawnSync(process.execPath, [copy], {
        cwd: directory,
        encoding: 'utf8',
        env: { ...process.env, NODE_PATH: '' },
        timeout: 30000
    });
    assert.ifError(result.error);
    assert.equal(result.status, 1);
    assert.match(result.stderr, /npm ci --ignore-scripts/);
    assert.match(result.stderr, /Cannot find module 'glob'/);
});

test('JSONC configuration preserves comment-like text inside strings', (t) => {
    const root = makeTempDir(t);
    writeFile(root, '.markdownlint.jsonc', [
        '{',
        '  // A real JSONC comment.',
        '  "MD013": false,',
        '  "example": "https://example.test/path//kept"',
        '}',
        ''
    ].join('\n'));
    const config = linter.loadMarkdownlintConfig(root);
    assert.equal(config.MD013, false);
    assert.equal(config.example, 'https://example.test/path//kept');
});

test('recursive extraction preserves nested Markdown fence depth', () => {
    const content = [
        '````markdown',
        '# Outer',
        '',
        '```markdown',
        '# Inner',
        '```',
        '````',
        ''
    ].join('\n');
    const blocks = linter.extractMarkdownFencesRecursive(content, 'fixture.md');
    assert.equal(blocks.length, 2);
    assert.deepEqual(blocks.map((block) => block.depth), [0, 1]);
    assert.deepEqual(blocks.map((block) => block.info), ['markdown', 'markdown']);
});

test('nested linting remains asynchronous and applies nested-only rules', async () => {
    let observed;
    const pending = linter.lintMarkdownContent('body\n', { MD013: false }, async (options) => {
        await Promise.resolve();
        observed = options;
        return { content: [] };
    });
    assert.equal(typeof pending.then, 'function');
    assert.deepEqual(await pending, { content: [] });
    assert.equal(observed.config.MD013, false);
    assert.equal(observed.config.MD041, false);
    assert.equal(observed.config.MD051, false);
});

test('regular files with spaces are canonicalized and readable no-block files succeed', (t) => {
    const parent = makeTempDir(t);
    const root = path.join(parent, 'repository with spaces');
    const input = writeFile(root, 'docs/file with spaces.md', '# Plain Markdown\n');
    const expected = fs.realpathSync(input);
    assert.equal(linter.validateMarkdownInput(root, input), expected);
    assert.deepEqual(linter.extractMarkdownFences(input, root), []);
});

test('read failures throw contextual errors and preserve their cause', (t) => {
    const root = makeTempDir(t);
    const input = writeFile(root, 'unreadable.md', '# Content\n');
    const cause = Object.assign(new Error('synthetic access denied'), { code: 'EACCES' });
    const fileSystem = {
        lstatSync: fs.lstatSync,
        realpathSync: fs.realpathSync,
        readFileSync() {
            throw cause;
        }
    };
    assert.throws(
        () => linter.extractMarkdownFences(input, root, fileSystem),
        (error) => {
            assert.equal(error.cause, cause);
            assert.match(error.message, /Could not read Markdown input/);
            assert.match(error.message, /unreadable\.md/);
            assert.match(error.message, /synthetic access denied/);
            return true;
        }
    );
});

test('rejected projected inputs are never read', () => {
    const root = path.resolve('projected-repository');
    const input = path.join(root, 'rules', 'candidate.mdc');
    const outside = path.resolve('projected-outside', 'outside.mdc');
    const sibling = path.resolve(`${root}-other`, 'sibling.mdc');
    const metadata = (symbolicLink, regularFile) => ({
        isSymbolicLink: () => symbolicLink,
        isFile: () => regularFile
    });
    const cases = [
        ['leaf symbolic link', metadata(true, true), outside, /non-symlink regular file/],
        ['directory', metadata(false, false), input, /non-symlink regular file/],
        ['outside target', metadata(false, true), outside, /outside the repository/],
        ['sibling-prefix target', metadata(false, true), sibling, /outside the repository/]
    ];

    for (const [label, inputMetadata, resolvedInput, expected] of cases) {
        let reads = 0;
        const fileSystem = {
            lstatSync: () => inputMetadata,
            realpathSync: (target) => path.resolve(target) === root ? root : resolvedInput,
            readFileSync: () => {
                reads += 1;
                return '';
            }
        };
        assert.throws(() => linter.extractMarkdownFences(input, root, fileSystem), expected, label);
        assert.equal(reads, 0, `${label} reached a content read.`);
    }
});

test('real leaf symbolic links are rejected before reads when supported', (t) => {
    const parent = makeTempDir(t);
    const root = path.join(parent, 'repository');
    const outside = writeFile(parent, 'outside.md', '# Outside\n');
    fs.mkdirSync(root);
    const link = path.join(root, 'linked.md');
    try {
        fs.symlinkSync(outside, link, 'file');
    } catch (error) {
        if (error.code === 'EPERM' || error.code === 'EACCES' || error.code === 'ENOTSUP') {
            t.skip(`File symbolic links are unavailable: ${error.code}`);
            return;
        }
        throw error;
    }
    assert.throws(() => linter.extractMarkdownFences(link, root), /non-symlink regular file/);
});

test('real parent-link escapes are rejected when directory links are supported', (t) => {
    const parent = makeTempDir(t);
    const root = path.join(parent, 'repository');
    const outside = path.join(parent, 'outside');
    fs.mkdirSync(root);
    writeFile(outside, 'escaped.md', '# Outside\n');
    const link = path.join(root, 'linked');
    const linkError = tryDirectoryLink(outside, link);
    if (linkError) {
        t.skip(`Directory links are unavailable: ${linkError.code}`);
        return;
    }
    assert.throws(
        () => linter.extractMarkdownFences(path.join(link, 'escaped.md'), root),
        /outside the repository/
    );
});

test('missing explicit files fail in the resolver and executable', (t) => {
    const root = makeTempDir(t);
    const missing = path.join(root, 'missing.md');
    assert.throws(() => linter.resolveFilePaths([missing], root), /ENOENT/);
    const result = spawnSync(process.execPath, [linterPath, missing], {
        cwd: repoRoot,
        encoding: 'utf8',
        timeout: 30000
    });
    assert.ifError(result.error);
    assert.equal(result.status, 1);
    assert.match(result.stderr, /ENOENT/);
    assert.match(result.stderr, /missing\.md/);
});

test('explicit outside, sibling-prefix, and nonregular inputs fail validation', (t) => {
    const parent = makeTempDir(t);
    const root = path.join(parent, 'repository');
    const outside = writeFile(parent, 'outside.md', '# Outside\n');
    const sibling = writeFile(`${root}-other`, 'sibling.md', '# Sibling\n');
    fs.mkdirSync(root);
    const directory = path.join(root, 'docs');
    fs.mkdirSync(directory);

    assert.throws(() => linter.resolveFilePaths([outside], root), /outside the repository/);
    assert.throws(() => linter.resolveFilePaths([sibling], root), /outside the repository/);
    assert.throws(() => linter.resolveFilePaths([directory], root), /non-symlink regular file/);

    const result = spawnSync(process.execPath, [linterPath, outside], {
        cwd: repoRoot,
        encoding: 'utf8',
        timeout: 30000
    });
    assert.ifError(result.error);
    assert.equal(result.status, 1);
    assert.match(result.stderr, /outside the repository/);
});

test('specific-file CLI accepts a repository file whose path contains spaces', (t) => {
    const isolated = makeIsolatedCliRepository(t, 'nested-markdown-cli-');
    const input = writeFile(isolated.root, 'docs/file with spaces.md', '# Plain Markdown\n');
    assert.equal(isWithin(repoRoot, isolated.root), false, 'CLI fixtures must stay outside the product repository.');
    const result = spawnSync(process.execPath, [isolated.executable, input], {
        cwd: isolated.root,
        encoding: 'utf8',
        env: { ...process.env, NODE_PATH: path.join(repoRoot, 'node_modules') },
        timeout: 30000
    });
    assert.ifError(result.error);
    assert.equal(result.status, 0, result.stderr);
    assert.match(result.stdout, /Linting 1 specified file/);
});

test('executable exits nonzero for a deterministic EACCES read failure', (t) => {
    const isolated = makeIsolatedCliRepository(t, 'nested-markdown-eacces-');
    const input = writeFile(isolated.root, 'fixtures/unreadable.md', '# Content\n');
    const preload = writeFile(isolated.root, 'fixtures/inject-eacces.cjs', [
        "const fs = require('fs');",
        "const path = require('path');",
        "const target = path.resolve(process.env.NESTED_MARKDOWN_UNREADABLE);",
        'const original = fs.readFileSync;',
        'fs.readFileSync = function(candidate, ...args) {',
        "    if (typeof candidate === 'string' && path.resolve(candidate) === target) {",
        "        const error = new Error('synthetic access denied');",
        "        error.code = 'EACCES';",
        '        throw error;',
        '    }',
        '    return original.call(this, candidate, ...args);',
        '};',
        ''
    ].join('\n'));
    assert.equal(isWithin(repoRoot, isolated.root), false, 'CLI fixtures must stay outside the product repository.');
    const result = spawnSync(process.execPath, ['--require', preload, isolated.executable, input], {
        cwd: isolated.root,
        encoding: 'utf8',
        env: {
            ...process.env,
            NESTED_MARKDOWN_UNREADABLE: input,
            NODE_PATH: path.join(repoRoot, 'node_modules')
        },
        timeout: 30000
    });
    assert.ifError(result.error);
    assert.equal(result.status, 1);
    assert.match(result.stderr, /Could not read Markdown input/);
    assert.match(result.stderr, /synthetic access denied/);
});

test('discovery includes hidden Markdown and mdc while pruning irrelevant trees', async (t) => {
    const root = makeTempDir(t);
    const included = [
        'README.md',
        'docs/visible.mdc',
        'docs/node_modules-guide.md',
        '.github/instructions/hidden.md',
        '.cursor/rules/hidden.mdc',
        '.hidden/nested.md'
    ];
    const excluded = [
        '.git/root.md',
        'nested/.git/nested.md',
        'node_modules/root.md',
        'nested/node_modules/nested.mdc',
        '.venv/root.md',
        'nested/.venv/nested.md',
        '.pytest_cache/root.md',
        'nested/.pytest_cache/nested.md',
        '.mypy_cache/root.md',
        'nested/.ruff_cache/nested.mdc',
        '__pycache__/root.md',
        'nested/__pycache__/nested.md'
    ];
    for (const file of [...included, ...excluded]) writeFile(root, file, '# Fixture\n');
    const discovered = await linter.findMarkdownFiles(root);
    assert.deepEqual(relativePaths(root, discovered), included.sort());
});

test('actual discovery does not follow a directory link outside the root', async (t) => {
    const parent = makeTempDir(t);
    const root = path.join(parent, 'repository');
    const outside = path.join(parent, 'outside');
    writeFile(root, 'visible.md', '# Visible\n');
    writeFile(outside, 'escaped.md', '# Outside\n');
    const linkError = tryDirectoryLink(outside, path.join(root, 'linked'));
    if (linkError) {
        t.skip(`Directory links are unavailable: ${linkError.code}`);
        return;
    }
    const discovered = await linter.findMarkdownFiles(root);
    assert.deepEqual(relativePaths(root, discovered), ['visible.md']);
});

test('actual discovery rejects a leaf symbolic link when supported', async (t) => {
    const parent = makeTempDir(t);
    const root = path.join(parent, 'repository');
    const outside = writeFile(parent, 'outside.md', '# Outside\n');
    fs.mkdirSync(root);
    const link = path.join(root, 'linked.md');
    try {
        fs.symlinkSync(outside, link, 'file');
    } catch (error) {
        if (error.code === 'EPERM' || error.code === 'EACCES' || error.code === 'ENOTSUP') {
            t.skip(`File symbolic links are unavailable: ${error.code}`);
            return;
        }
        throw error;
    }
    await assert.rejects(linter.findMarkdownFiles(root), /non-symlink regular file/);
});

test('every injected discovery result passes through the input boundary', async () => {
    const root = path.resolve('projected-discovery-root');
    const candidate = path.join(root, 'linked.mdc');
    let captured;
    const globFunction = async (pattern, options) => {
        captured = { pattern, options };
        return [candidate];
    };
    const fileSystem = {
        lstatSync: () => ({ isFile: () => true, isSymbolicLink: () => true }),
        realpathSync: (target) => path.resolve(target)
    };
    await assert.rejects(
        linter.findMarkdownFiles(root, globFunction, fileSystem),
        /non-symlink regular file/
    );
    assert.equal(captured.pattern, '**/*.{md,mdc}');
    assert.equal(captured.options.dot, true);
    assert.equal(captured.options.follow, false);
    assert.equal(captured.options.nodir, true);
    assert.equal(captured.options.absolute, true);
});

test('read-failure mutant restoring the swallowed result is detected', (t) => {
    const mutant = createMutant(
        t,
        /throw new Error\(\s*`Could not read Markdown input \$\{safeInputPath\}: \$\{error\.message\}`,\s*\{ cause: error \}\s*\);/,
        'return [];',
        'swallowed-read'
    );
    const script = [
        "const assert = require('assert/strict');",
        "const path = require('path');",
        'const subject = require(process.argv[1]);',
        "const root = path.resolve('root');",
        "const input = path.join(root, 'unreadable.md');",
        'const fileSystem = {',
        '  lstatSync: () => ({ isFile: () => true, isSymbolicLink: () => false }),',
        '  realpathSync: (target) => path.resolve(target),',
        "  readFileSync: () => { const error = new Error('synthetic'); error.code = 'EACCES'; throw error; }",
        '};',
        'assert.throws(() => subject.extractMarkdownFences(input, root, fileSystem), /Could not read/);'
    ].join('\n');
    assertMutantDetected(runMutant(mutant, script), 'swallowed read failure');
});

test('symlink-guard removal mutant is detected independently', (t) => {
    const mutant = createMutant(
        t,
        'inputMetadata.isSymbolicLink() || !inputMetadata.isFile()',
        '!inputMetadata.isFile()',
        'symlink-guard'
    );
    const script = [
        "const assert = require('assert/strict');",
        "const path = require('path');",
        'const subject = require(process.argv[1]);',
        "const root = path.resolve('root');",
        "const input = path.join(root, 'linked.md');",
        'let reads = 0;',
        'const fileSystem = {',
        '  lstatSync: () => ({ isFile: () => true, isSymbolicLink: () => true }),',
        '  realpathSync: (target) => path.resolve(target),',
        "  readFileSync: () => { reads += 1; return ''; }",
        '};',
        'assert.throws(() => subject.extractMarkdownFences(input, root, fileSystem), /non-symlink/);',
        "assert.equal(reads, 0, 'Rejected symlink reached a content read.');"
    ].join('\n');
    assertMutantDetected(runMutant(mutant, script), 'symlink guard removal');
});

test('containment-guard removal mutant is detected independently', (t) => {
    const mutant = createMutant(
        t,
        /if \(relativeInputPath === '\.\.' \|\|\s*relativeInputPath\.startsWith\(`\.\.\$\{path\.sep\}`\) \|\|\s*path\.isAbsolute\(relativeInputPath\)\) \{/,
        'if (false) {',
        'containment-guard'
    );
    const script = [
        "const assert = require('assert/strict');",
        "const path = require('path');",
        'const subject = require(process.argv[1]);',
        "const root = path.resolve('root');",
        "const input = path.join(root, 'candidate.md');",
        "const outside = path.resolve('outside', 'candidate.md');",
        'let reads = 0;',
        'const fileSystem = {',
        '  lstatSync: () => ({ isFile: () => true, isSymbolicLink: () => false }),',
        '  realpathSync: (target) => path.resolve(target) === root ? root : outside,',
        "  readFileSync: () => { reads += 1; return ''; }",
        '};',
        'assert.throws(() => subject.extractMarkdownFences(input, root, fileSystem), /outside/);',
        "assert.equal(reads, 0, 'Rejected outside input reached a content read.');"
    ].join('\n');
    assertMutantDetected(runMutant(mutant, script), 'containment guard removal');
});

test('hidden-file discovery mutant is detected independently', (t) => {
    const root = makeTempDir(t);
    writeFile(root, '.github/instructions/hidden.md', '# Hidden\n');
    const mutant = createMutant(t, 'dot: true', 'dot: false', 'hidden-discovery');
    const script = [
        "const assert = require('assert/strict');",
        "const path = require('path');",
        'const subject = require(process.argv[1]);',
        'subject.findMarkdownFiles(process.argv[2]).then((files) => {',
        "  assert(files.some((file) => path.basename(file) === 'hidden.md'));",
        '}).catch((error) => { console.error(error); process.exitCode = 1; });'
    ].join('\n');
    assertMutantDetected(runMutant(mutant, script, [root]), 'hidden-file discovery');
});

test('mdc discovery mutant is detected independently', (t) => {
    const root = makeTempDir(t);
    writeFile(root, 'rules/example.mdc', '# Rule\n');
    const mutant = createMutant(t, "'**/*.{md,mdc}'", "'**/*.md'", 'mdc-discovery');
    const script = [
        "const assert = require('assert/strict');",
        "const path = require('path');",
        'const subject = require(process.argv[1]);',
        'subject.findMarkdownFiles(process.argv[2]).then((files) => {',
        "  assert(files.some((file) => path.extname(file) === '.mdc'));",
        '}).catch((error) => { console.error(error); process.exitCode = 1; });'
    ].join('\n');
    assertMutantDetected(runMutant(mutant, script, [root]), 'mdc discovery');
});

test('Git exclusion mutant is detected independently', (t) => {
    const root = makeTempDir(t);
    writeFile(root, '.git/ignored.md', '# Ignored\n');
    writeFile(root, 'visible.md', '# Visible\n');
    const mutant = createMutant(
        t,
        /\s*'\.git\/\*\*',\s*'\*\*\/\.git\/\*\*',/,
        '',
        'git-exclusion'
    );
    const script = [
        "const assert = require('assert/strict');",
        "const path = require('path');",
        'const subject = require(process.argv[1]);',
        'subject.findMarkdownFiles(process.argv[2]).then((files) => {',
        "  const relative = files.map((file) => path.relative(process.argv[2], file).split(path.sep).join('/'));",
        "  assert.deepEqual(relative.sort(), ['visible.md']);",
        '}).catch((error) => { console.error(error); process.exitCode = 1; });'
    ].join('\n');
    assertMutantDetected(runMutant(mutant, script, [root]), 'Git exclusion');
});

test('no-follow option mutant is detected independently', (t) => {
    const mutant = createMutant(t, 'follow: false', 'follow: true', 'no-follow');
    const script = [
        "const assert = require('assert/strict');",
        "const path = require('path');",
        'const subject = require(process.argv[1]);',
        "const root = path.resolve('root');",
        'const globFunction = async (pattern, options) => {',
        "  assert.equal(pattern, '**/*.{md,mdc}');",
        "  assert.equal(options.follow, false, 'Discovery must not follow directory links.');",
        '  return [];',
        '};',
        'subject.findMarkdownFiles(root, globFunction).catch((error) => {',
        '  console.error(error);',
        '  process.exitCode = 1;',
        '});'
    ].join('\n');
    assertMutantDetected(runMutant(mutant, script), 'no-follow option');
});
