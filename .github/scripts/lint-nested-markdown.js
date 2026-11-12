#!/usr/bin/env node

/**
 * Lint Nested Markdown Script
 *
 * This script extracts Markdown code blocks from Markdown files and runs
 * markdownlint on them to ensure nested Markdown content follows the same
 * linting rules as the outer Markdown files.
 *
 * Usage:
 *   Scan all files:   node .github/scripts/lint-nested-markdown.js
 *   Lint specific files:  node .github/scripts/lint-nested-markdown.js file1.md file2.md
 *
 * When file arguments are provided, only those files are linted (useful for pre-commit hooks).
 * When no arguments are provided, all .md and .mdc files are scanned via glob.
 * Both absolute and relative paths are supported; relative paths are resolved from cwd.
 */

const fs = require('fs');
const path = require('path');

// Repository root is two levels up from this script's location in .github/scripts/
const REPO_ROOT = path.resolve(__dirname, '../..');

const MARKDOWN_GLOB = '**/*.{md,mdc}';
// Per-file budgets bound repeated parsing and linting of overlapping nested bodies.
const MAX_MARKDOWN_BYTES = 1024 * 1024;
const MAX_NESTING_DEPTH = 64;
const MAX_EXTRACTED_BLOCKS = 1024;
const MAX_TOTAL_EXTRACTED_BYTES = 8 * 1024 * 1024;
const MARKDOWN_IGNORE = [
    'node_modules/**',
    '**/node_modules/**',
    '.git/**',
    '**/.git/**',
    '.venv/**',
    '**/.venv/**',
    '.pytest_cache/**',
    '**/.pytest_cache/**',
    '.mypy_cache/**',
    '**/.mypy_cache/**',
    '.ruff_cache/**',
    '**/.ruff_cache/**',
    '__pycache__/**',
    '**/__pycache__/**'
];

/**
 * Load the external packages required by the nested-Markdown linter.
 * @param {NodeRequire} moduleLoader - CommonJS module loader
 * @param {string} repoRoot - Repository root used in setup guidance
 * @returns {object} Loaded runtime dependencies
 */
function loadRuntimeDependencies(moduleLoader = require, repoRoot = REPO_ROOT) {
    try {
        const { glob } = moduleLoader('glob');
        const MarkdownIt = moduleLoader('markdown-it');
        const { lint } = moduleLoader('markdownlint/promise');
        const jsoncParser = moduleLoader('jsonc-parser');
        return { glob, MarkdownIt, lint, jsoncParser };
    } catch (error) {
        if (error.code !== 'MODULE_NOT_FOUND') {
            throw error;
        }
        throw new Error(
            `Unable to load nested-Markdown dependencies. Run ` +
            `\`npm ci --ignore-scripts\` from ${repoRoot}. Cause: ${error.message}`,
            { cause: error }
        );
    }
}

const { glob, MarkdownIt, lint, jsoncParser } = loadRuntimeDependencies();

// Initialize markdown-it parser
const md = new MarkdownIt();

// ANSI color codes for terminal output
const colors = {
    reset: '\x1b[0m',
    red: '\x1b[31m',
    yellow: '\x1b[33m',
    green: '\x1b[32m',
    cyan: '\x1b[36m',
    bold: '\x1b[1m'
};

/**
 * Resolve file paths from command-line arguments to absolute paths
 * @param {string[]} args - Command-line arguments (file paths)
 * @param {string} repoRoot - Repository root
 * @param {object} fileSystem - File-system adapter used by deterministic tests
 * @returns {string[]} Canonical validated input paths
 */
function resolveFilePaths(args, repoRoot = REPO_ROOT, fileSystem = fs) {
    const validFiles = [];

    for (const arg of args) {
        // Resolve relative paths from current working directory
        const absolutePath = path.isAbsolute(arg)
            ? arg
            : path.resolve(process.cwd(), arg);
        validFiles.push(validateMarkdownInput(repoRoot, absolutePath, fileSystem));
    }

    return validFiles;
}

/**
 * Validate one nested-Markdown input before reading it.
 * @param {string} repoRoot - Repository root
 * @param {string} filePath - Candidate Markdown input
 * @param {object} fileSystem - File-system adapter used by deterministic tests
 * @returns {string} Canonical in-repository input path
 */
function validateMarkdownInput(repoRoot, filePath, fileSystem = fs) {
    const rootPath = fileSystem.realpathSync(repoRoot);
    const absoluteInputPath = path.resolve(filePath);
    const inputMetadata = fileSystem.lstatSync(absoluteInputPath);

    if (inputMetadata.isSymbolicLink() || !inputMetadata.isFile()) {
        throw new Error(`Markdown input must be a non-symlink regular file: ${filePath}`);
    }

    const resolvedInputPath = fileSystem.realpathSync(absoluteInputPath);
    const relativeInputPath = path.relative(rootPath, resolvedInputPath);
    if (relativeInputPath === '..' ||
        relativeInputPath.startsWith(`..${path.sep}`) ||
        path.isAbsolute(relativeInputPath)) {
        throw new Error(`Markdown input resolves outside the repository: ${filePath}`);
    }

    return resolvedInputPath;
}

/**
 * Discover and validate Markdown inputs below a repository root.
 * @param {string} repoRoot - Repository root
 * @param {Function} globFunction - Glob adapter used by deterministic tests
 * @param {object} fileSystem - File-system adapter used by deterministic tests
 * @returns {Promise<string[]>} Canonical validated input paths
 */
async function findMarkdownFiles(repoRoot = REPO_ROOT, globFunction = glob, fileSystem = fs) {
    const files = await globFunction(MARKDOWN_GLOB, {
        ignore: MARKDOWN_IGNORE,
        cwd: repoRoot,
        dot: true,
        absolute: true,
        follow: false,
        nodir: true
    });
    return files.map((file) => validateMarkdownInput(repoRoot, file, fileSystem));
}

/**
 * Load markdownlint configuration from .markdownlint.jsonc or .markdownlint.json
 */
function loadMarkdownlintConfig(repoRoot = REPO_ROOT, fileSystem = fs, parser = jsoncParser) {
    const configPaths = [
        path.join(repoRoot, '.markdownlint.jsonc'),
        path.join(repoRoot, '.markdownlint.json')
    ];

    for (const configPath of configPaths) {
        if (fileSystem.existsSync(configPath)) {
            try {
                const content = fileSystem.readFileSync(configPath, 'utf8');
                // Use jsonc-parser for proper JSONC handling (supports comments in strings)
                return parser.parse(content);
            } catch (error) {
                console.warn(`Warning: Could not read config file ${configPath}: ${error.message}`);
                continue;
            }
        }
    }
    return {};
}

/**
 * Extract markdown code fences from content (recursive)
 * @param {string} content - Markdown content to parse
 * @param {string} filePath - Path to the original markdown file
 * @param {number} baseLine - Line number offset in the original file
 * @param {number} depth - Current nesting depth
 * @param {string} parentPath - Path description for nested blocks
 * @param {object} budget - Shared counters for the entire original input
 * @returns {Array} Array of extracted blocks with metadata
 */
function extractMarkdownFencesRecursive(content, filePath, baseLine = 0, depth = 0, parentPath = '', budget = { blocks: 0, bytes: 0 }) {
    if (Buffer.byteLength(content, 'utf8') > MAX_MARKDOWN_BYTES) {
        throw new Error(`Markdown budget exceeded in ${filePath}: source bytes limit ${MAX_MARKDOWN_BYTES}`);
    }
    const tokens = md.parse(content, {});
    const blocks = [];

    for (let i = 0; i < tokens.length; i++) {
        const token = tokens[i];

        // Look for fence tokens with markdown language identifier
        if (token.type === 'fence' &&
            (token.info.trim().toLowerCase() === 'markdown' ||
             token.info.trim().toLowerCase() === 'md')) {

            if (depth >= MAX_NESTING_DEPTH) {
                throw new Error(`Markdown budget exceeded in ${filePath}: nesting depth limit ${MAX_NESTING_DEPTH}`);
            }
            budget.blocks += 1;
            budget.bytes += Buffer.byteLength(token.content, 'utf8');
            if (budget.blocks > MAX_EXTRACTED_BLOCKS) {
                throw new Error(`Markdown budget exceeded in ${filePath}: block count limit ${MAX_EXTRACTED_BLOCKS}`);
            }
            if (budget.bytes > MAX_TOTAL_EXTRACTED_BYTES) {
                throw new Error(`Markdown budget exceeded in ${filePath}: extracted bytes limit ${MAX_TOTAL_EXTRACTED_BYTES}`);
            }

            const blockLine = baseLine + (token.map ? token.map[0] + 1 : 0);
            const blockPath = parentPath ? `${parentPath} > block at line ${blockLine}` : `line ${blockLine}`;

            const blockInfo = {
                content: token.content,
                line: blockLine,
                info: token.info.trim(),
                filePath: filePath,
                depth: depth,
                parentPath: blockPath
            };

            blocks.push(blockInfo);

            // Recursively extract nested markdown fences
            if (token.content.trim().length > 0) {
                const nestedBlocks = extractMarkdownFencesRecursive(
                    token.content,
                    filePath,
                    blockLine,
                    depth + 1,
                    blockPath,
                    budget
                );
                blocks.push(...nestedBlocks);
            }
        }
    }

    return blocks;
}

/**
 * Extract markdown code fences from a file
 * @param {string} filePath - Path to the markdown file
 * @param {string} repoRoot - Repository root
 * @param {object} fileSystem - File-system adapter used by deterministic tests
 * @returns {Array} Array of extracted blocks with metadata
 */
function extractMarkdownFences(filePath, repoRoot = REPO_ROOT, fileSystem = fs) {
    const safeInputPath = validateMarkdownInput(repoRoot, filePath, fileSystem);
    if (fileSystem.lstatSync(safeInputPath).size > MAX_MARKDOWN_BYTES) {
        throw new Error(`Markdown budget exceeded in ${filePath}: source bytes limit ${MAX_MARKDOWN_BYTES}`);
    }
    let content;
    try {
        content = fileSystem.readFileSync(safeInputPath, 'utf8');
    } catch (error) {
        throw new Error(
            `Could not read Markdown input ${safeInputPath}: ${error.message}`,
            { cause: error }
        );
    }
    return extractMarkdownFencesRecursive(content, safeInputPath, 0, 0, '');
}

/**
 * Run markdownlint on extracted content
 * @param {string} content - Markdown content to lint
 * @param {object} config - Markdownlint configuration
 * @param {Function} lintFunction - Asynchronous markdownlint adapter
 * @returns {Promise<object>} Markdownlint results
 */
async function lintMarkdownContent(content, config, lintFunction = lint) {
    // Create a modified config for nested markdown
    // Disable MD041 (first-line-heading) since nested markdown snippets
    // may not start with a top-level heading
    // Disable MD051 (link-fragments) since nested markdown often contains
    // example/placeholder links that reference anchors in other documents
    const nestedConfig = {
        ...config,
        'MD041': false,
        'MD051': false
    };

    const options = {
        strings: {
            'content': content
        },
        config: nestedConfig
    };

    return await lintFunction(options);
}

/**
 * Format and display linting results
 * @param {Array} allResults - Array of results with context
 * @returns {boolean} True if any errors were found
 */
function displayResults(allResults) {
    let hasErrors = false;

    if (allResults.length === 0) {
        console.log(`${colors.green}✓${colors.reset} No issues found in nested Markdown code fences`);
        return false;
    }

    console.log(`\n${colors.bold}${colors.red}Nested Markdown Linting Issues:${colors.reset}\n`);

    for (const result of allResults) {
        if (result.errors.length === 0) {
            continue;
        }

        hasErrors = true;

        console.log(`${colors.cyan}File:${colors.reset} ${result.filePath}`);
        const depthIndicator = result.depth > 0 ? ` ${colors.yellow}[depth ${result.depth}]${colors.reset}` : '';
        const pathInfo = result.parentPath ? ` (${result.parentPath})` : '';
        console.log(`  ${colors.yellow}Code fence at line ${result.line}${depthIndicator} (${result.info} block #${result.blockIndex})${pathInfo}:${colors.reset}`);

        for (const error of result.errors) {
            // Calculate the actual line number in the outer file
            // result.line is the fence opening line (e.g., line 9)
            // error.lineNumber is 1-based line within the content (e.g., line 1 is first content line)
            // Content starts at result.line + 1, so line N of content is at result.line + N
            const actualLineNumber = result.line + error.lineNumber;
            const nestedLineInfo = result.depth > 0 ? ` (nested line ${error.lineNumber})` : '';
            console.log(`    ${actualLineNumber}:${error.errorRange ? error.errorRange[0] : 1}${nestedLineInfo} ${colors.red}${error.ruleNames.join('/')}${colors.reset} ${error.ruleDescription}`);
            if (error.errorDetail) {
                console.log(`      ${colors.yellow}${error.errorDetail}${colors.reset}`);
            }
        }

        console.log('');
    }

    return hasErrors;
}

/**
 * Main function
 */
async function main() {
    try {
        console.log(`${colors.bold}Linting nested Markdown in code fences...${colors.reset}\n`);

        // Load markdownlint configuration
        const config = loadMarkdownlintConfig();

        // Check for command-line file arguments
        const cliArgs = process.argv.slice(2);
        let files;

        if (cliArgs.length > 0) {
            // Use files provided as arguments
            files = resolveFilePaths(cliArgs);
            console.log(`Linting ${files.length} specified file(s)\n`);
        } else {
            // Find all markdown files (excluding generated/dependency directories)
            files = await findMarkdownFiles();
            console.log(`Found ${files.length} Markdown file(s) to scan\n`);
        }

        let totalBlocks = 0;
        const allResults = [];

        // Process each file
        for (const file of files) {
            const relativePath = path.relative(REPO_ROOT, file);
            const blocks = extractMarkdownFences(file);

            if (blocks.length > 0) {
                console.log(`${colors.cyan}${relativePath}${colors.reset}: Found ${blocks.length} nested Markdown block(s)`);
                totalBlocks += blocks.length;

                // Lint each extracted block
                for (let index = 0; index < blocks.length; index++) {
                    const block = blocks[index];

                    // Skip empty blocks
                    if (!block.content || block.content.trim().length === 0) {
                        continue;
                    }

                    const lintResults = await lintMarkdownContent(block.content, config);
                    const errors = lintResults.content || [];

                    if (errors.length > 0) {
                        allResults.push({
                            filePath: relativePath,
                            line: block.line,
                            info: block.info,
                            blockIndex: index + 1,
                            depth: block.depth,
                            parentPath: block.parentPath,
                            errors: errors
                        });
                    }
                }
            }
        }

        console.log(`\nTotal nested Markdown blocks found: ${totalBlocks}\n`);

        // Display results
        const hasErrors = displayResults(allResults);

        if (hasErrors) {
            console.log(`${colors.red}${colors.bold}✗${colors.reset} ${colors.red}Nested Markdown linting failed${colors.reset}\n`);
            process.exit(1);
        } else {
            console.log(`${colors.green}${colors.bold}✓${colors.reset} ${colors.green}Nested Markdown linting passed${colors.reset}\n`);
            process.exit(0);
        }

    } catch (error) {
        console.error(`${colors.red}Error:${colors.reset}`, error.message);
        console.error(error.stack);
        process.exit(1);
    }
}

// Run main function only for direct CLI execution.
if (require.main === module) {
    main();
}

module.exports = {
    extractMarkdownFences,
    extractMarkdownFencesRecursive,
    findMarkdownFiles,
    lintMarkdownContent,
    loadMarkdownlintConfig,
    loadRuntimeDependencies,
    main,
    resolveFilePaths,
    validateMarkdownInput
};
