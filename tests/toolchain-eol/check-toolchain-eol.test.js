const assert = require('assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawnSync } = require('child_process');
const { test, after } = require('node:test');
const yaml = require('yaml');

const scanner = require('../../.github/scripts/check-toolchain-eol.js');

const FIXTURE_SCHEDULE = require('./fixtures/node-schedule.json');

const EXPECTED_TOOLCHAIN_EOL_TEST_SCRIPT = 'node --test tests/toolchain-eol/check-toolchain-eol.test.js';
const createdTempRepos = [];

function makeTempRepo() {
    const repoRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'toolchain-eol-'));
    createdTempRepos.push(repoRoot);
    return repoRoot;
}

after(() => {
    for (const repoRoot of createdTempRepos) {
        fs.rmSync(repoRoot, { recursive: true, force: true });
    }
});

function writeFile(repoRoot, relativePath, content) {
    const filePath = path.join(repoRoot, relativePath);
    fs.mkdirSync(path.dirname(filePath), { recursive: true });
    fs.writeFileSync(filePath, content, 'utf8');
}

function selectorKey(selector) {
    const referenced = selector.referencedPath ? ` via ${selector.referencedPath}` : '';
    return `${selector.selectorClass}|${selector.sourceType}|${selector.origin}|${selector.rawValue}|${selector.path}${referenced}`;
}

test('package script uses the explicit toolchain EOL test file target', () => {
    const packageJsonPath = path.resolve(__dirname, '..', '..', 'package.json');
    const packageJson = JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));

    assert.equal(
        packageJson.scripts['test:toolchain-eol'],
        EXPECTED_TOOLCHAIN_EOL_TEST_SCRIPT,
        [
            'scripts["test:toolchain-eol"] must use the selected explicit-file target',
            `"${EXPECTED_TOOLCHAIN_EOL_TEST_SCRIPT}" and must not be reverted to the bare-directory`,
            '"node --test tests/toolchain-eol/" target; directory recursion was documented in Node v20',
            'and reported as no longer working in Node v21.',
        ].join(' '),
    );
});

test('explicit toolchain EOL npm target covers every test file in this directory', () => {
    const testFiles = fs.readdirSync(__dirname).filter((name) => name.endsWith('.test.js')).sort();

    assert.deepEqual(
        testFiles,
        ['check-toolchain-eol.test.js'],
        [
            'The explicit test:toolchain-eol npm target must cover every tests/toolchain-eol/*.test.js file.',
            'If a new *.test.js file is added, add it to the npm target and this assertion, and review or',
            'update .template-sync/manifest.yml so the new test file has the correct module ownership.',
        ].join(' '),
    );
});

test('discovers and classifies checked-in Node.js selectors without action-version false positives', () => {
    const repoRoot = makeTempRepo();
    writeFile(
        repoRoot,
        'package.json',
        JSON.stringify({ engines: { node: '>=22.0.0' } }, null, 2),
    );
    writeFile(
        repoRoot,
        'package-lock.json',
        JSON.stringify(
            {
                packages: {
                    '': { engines: { node: '>=22.0.0' } },
                    'node_modules/transitive': { engines: { node: '>=10' } },
                },
            },
            null,
            2,
        ),
    );
    writeFile(repoRoot, '.nvmrc', '23\n');
    writeFile(repoRoot, '.node-version', '24\n');

    writeFile(
        repoRoot,
        '.github/workflows/literal.yml',
        `
name: Literal
on:
  workflow_dispatch:
jobs:
  test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/setup-node@v6
        with:
          node-version: "24"
`,
    );
    writeFile(
        repoRoot,
        '.github/workflows/matrix.yml',
        `
name: Matrix
on:
  workflow_dispatch:
jobs:
  test:
    runs-on: ubuntu-latest
    strategy:
      matrix:
        node-version: ["22", "24"]
    steps:
      - uses: actions/setup-node@v6
        with:
          node-version: \${{ matrix.node-version }}
`,
    );
    writeFile(
        repoRoot,
        '.github/workflows/version-file.yml',
        `
name: Version file
on:
  workflow_dispatch:
jobs:
  test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/setup-node@v6
        with:
          node-version-file: ".nvmrc"
`,
    );

    writeFile(
        repoRoot,
        '.azuredevops/pipelines/parameters.yml',
        `
parameters:
  - name: nodeVersion
    type: string
    default: "22"
    values:
      - "22"
      - "24"
steps:
  - task: UseNode@1
    inputs:
      version: \${{ parameters.nodeVersion }}
`,
    );
    writeFile(
        repoRoot,
        '.azuredevops/pipelines/variables.yml',
        `
variables:
  nodeVersion: "22"
steps:
  - task: UseNode@1
    inputs:
      version: "$(nodeVersion)"
`,
    );
    writeFile(
        repoRoot,
        '.azuredevops/pipelines/stage-and-job-variables.yml',
        `
stages:
  - stage: build
    variables:
      stageNodeVersion: "22"
    jobs:
      - job: stageVariable
        steps:
          - task: UseNode@1
            inputs:
              version: "$(stageNodeVersion)"
      - job: jobVariable
        variables:
          jobNodeVersion: "24"
        steps:
          - task: UseNode@1
            inputs:
              version: "$(jobNodeVersion)"
`,
    );
    writeFile(
        repoRoot,
        '.azuredevops/pipelines/matrix.yml',
        `
jobs:
  - job: test
    strategy:
      matrix:
        node22:
          nodeVersion: "22"
        node24:
          nodeVersion: "24"
    steps:
      - task: UseNode@1
        inputs:
          version: "$(nodeVersion)"
`,
    );
    writeFile(
        repoRoot,
        '.azuredevops/pipelines/nodetool.yml',
        `
steps:
  - task: NodeTool@0
    inputs:
      versionSpec: "23"
`,
    );
    writeFile(
        repoRoot,
        '.azuredevops/pipelines/from-file.yml',
        `
steps:
  - task: NodeTool@0
    inputs:
      versionSource: "fromFile"
      versionFilePath: ".node-version"
`,
    );
    writeFile(
        repoRoot,
        '.azuredevops/pipelines/use-node.yml',
        `
steps:
  - task: UseNode@1
    inputs:
      version: "24"
`,
    );

    const inventory = scanner.collectNodeSelectors(repoRoot);
    assert.deepEqual(inventory.problems, []);

    const keys = inventory.selectors.map(selectorKey);
    assert(keys.includes('support-floor|package-json:engines.node|root package engines.node|>=22.0.0|package.json'));
    assert(keys.includes('support-floor|package-lock:root.engines.node|root package-lock mirror|>=22.0.0|package-lock.json'));
    assert(keys.includes('ci-runtime|github-actions:setup-node node-version|literal|24|.github/workflows/literal.yml'));
    assert(keys.includes('ci-runtime|github-actions:setup-node node-version|matrix.node-version|22|.github/workflows/matrix.yml'));
    assert(keys.includes('ci-runtime|github-actions:setup-node node-version|matrix.node-version|24|.github/workflows/matrix.yml'));
    assert(keys.includes('ci-runtime|github-actions:setup-node node-version-file|version-file|23|.github/workflows/version-file.yml via .nvmrc'));
    assert(keys.includes('ci-runtime|azure-pipelines:UseNode@1 version|parameters.nodeVersion default-or-values|22|.azuredevops/pipelines/parameters.yml'));
    assert(keys.includes('ci-runtime|azure-pipelines:UseNode@1 version|parameters.nodeVersion default-or-values|24|.azuredevops/pipelines/parameters.yml'));
    assert(keys.includes('ci-runtime|azure-pipelines:UseNode@1 version|variable-or-matrix.nodeVersion|22|.azuredevops/pipelines/variables.yml'));
    assert(keys.includes('ci-runtime|azure-pipelines:UseNode@1 version|variable-or-matrix.stageNodeVersion|22|.azuredevops/pipelines/stage-and-job-variables.yml'));
    assert(keys.includes('ci-runtime|azure-pipelines:UseNode@1 version|variable-or-matrix.jobNodeVersion|24|.azuredevops/pipelines/stage-and-job-variables.yml'));
    assert(keys.includes('ci-runtime|azure-pipelines:UseNode@1 version|variable-or-matrix.nodeVersion|24|.azuredevops/pipelines/matrix.yml'));
    assert(keys.includes('ci-runtime|azure-pipelines:NodeTool@0 versionSpec|literal|23|.azuredevops/pipelines/nodetool.yml'));
    assert(keys.includes('ci-runtime|azure-pipelines:NodeTool@0 versionFilePath|version-file|24|.azuredevops/pipelines/from-file.yml via .node-version'));
    assert(keys.includes('ci-runtime|azure-pipelines:UseNode@1 version|literal|24|.azuredevops/pipelines/use-node.yml'));
    assert(!keys.some((key) => key.includes('setup-node@v6')));
    assert(!keys.some((key) => key.includes('>=10')));
});

test('evaluates support-floor ranges and EOL warning-window status', () => {
    const selectors = [
        {
            toolchain: 'node',
            selectorClass: 'support-floor',
            sourceType: 'package-json:engines.node',
            origin: 'root package engines.node',
            path: 'package.json',
            rawValue: '>=22.0.0',
        },
        {
            toolchain: 'node',
            selectorClass: 'ci-runtime',
            sourceType: 'github-actions:setup-node node-version',
            origin: 'literal',
            path: '.github/workflows/eol.yml',
            rawValue: '23',
        },
        {
            toolchain: 'node',
            selectorClass: 'ci-runtime',
            sourceType: 'github-actions:setup-node node-version',
            origin: 'literal',
            path: '.github/workflows/supported.yml',
            rawValue: '24',
        },
    ];

    const result = scanner.evaluateSelectors(selectors, FIXTURE_SCHEDULE, {
        asOfDate: '2026-01-15',
        warningWindowDays: 180,
    });

    assert.deepEqual(result.problems, []);

    const statusByPath = new Map(result.findings.map((finding) => [finding.path, finding]));
    assert.equal(statusByPath.get('package.json').releaseLine, 22);
    assert.equal(statusByPath.get('package.json').status, 'near-eol');
    assert.equal(statusByPath.get('.github/workflows/eol.yml').status, 'eol');
    assert.equal(statusByPath.get('.github/workflows/supported.yml').status, 'supported');
});

test('reports package-lock root engine drift without reading transitive engines as policy', () => {
    const repoRoot = makeTempRepo();
    writeFile(
        repoRoot,
        'package.json',
        JSON.stringify({ engines: { node: '>=22.0.0' } }, null, 2),
    );
    writeFile(
        repoRoot,
        'package-lock.json',
        JSON.stringify(
            {
                packages: {
                    '': { engines: { node: '>=24.0.0' } },
                    'node_modules/transitive': { engines: { node: '>=8' } },
                },
            },
            null,
            2,
        ),
    );

    const inventory = scanner.collectNodeSelectors(repoRoot);

    assert.equal(inventory.problems.length, 1);
    assert.match(inventory.problems[0].message, /does not match/);
    assert.equal(
        inventory.selectors.filter((selector) => selector.sourceType === 'package-lock:root.engines.node')
            .length,
        1,
    );
    assert(!inventory.selectors.some((selector) => selector.rawValue === '>=8'));
});

test('handles unquoted numeric GitHub Actions node-version selectors', () => {
    const repoRoot = makeTempRepo();
    writeFile(
        repoRoot,
        '.github/workflows/numeric.yml',
        `
name: Numeric
on:
  workflow_dispatch:
jobs:
  test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/setup-node@v6
        with:
          node-version: 24
`,
    );

    const inventory = scanner.collectNodeSelectors(repoRoot);
    assert.deepEqual(inventory.problems, []);

    const keys = inventory.selectors.map(selectorKey);
    assert(keys.includes('ci-runtime|github-actions:setup-node node-version|literal|24|.github/workflows/numeric.yml'));
    assert(!inventory.selectors.some((selector) => selector.rawValue === 'undefined'));

    // A non-string selector must still evaluate cleanly rather than cascading
    // into a false "unable to parse" inventory problem.
    const evaluation = scanner.evaluateSelectors(inventory.selectors, FIXTURE_SCHEDULE, {
        asOfDate: '2026-01-15',
        warningWindowDays: 180,
    });
    assert.deepEqual(evaluation.problems, []);
    assert.equal(evaluation.findings[0].releaseLine, 24);
});

test('escapes percent before CR/LF in GitHub Actions annotation data', () => {
    // A literal "%0A" in the text must not be decodable as a newline by the
    // runner command processor: the percent is escaped first, yielding "%250A".
    assert.equal(scanner.escapeWorkflowData('before%0Aafter'), 'before%250Aafter');
    // Real control characters are still encoded.
    assert.equal(scanner.escapeWorkflowData('a\r\nb'), 'a%0D%0Ab');
    // A bare percent is encoded.
    assert.equal(scanner.escapeWorkflowData('100%'), '100%25');
});

test('classifies a selector evaluated on its EOL date as eol', () => {
    const selectors = [
        {
            toolchain: 'node',
            selectorClass: 'ci-runtime',
            sourceType: 'github-actions:setup-node node-version',
            origin: 'literal',
            path: '.github/workflows/on-eol-date.yml',
            rawValue: '22',
        },
    ];

    // FIXTURE_SCHEDULE v22 end is 2026-07-01; evaluating on that exact date
    // means daysUntilEol === 0, which must be the stronger "eol" signal.
    const result = scanner.evaluateSelectors(selectors, FIXTURE_SCHEDULE, {
        asOfDate: '2026-07-01',
        warningWindowDays: 180,
    });

    assert.deepEqual(result.problems, []);
    assert.equal(result.findings[0].daysUntilEol, 0);
    assert.equal(result.findings[0].status, 'eol');
});

test('aborts the schedule fetch after the configured timeout', async () => {
    const originalFetch = global.fetch;
    // Simulate a stalled endpoint: the promise never settles on its own but
    // honors the abort signal the way the real fetch does.
    global.fetch = (url, opts) =>
        new Promise((resolve, reject) => {
            opts.signal.addEventListener('abort', () => {
                const abortError = new Error('aborted');
                abortError.name = 'AbortError';
                reject(abortError);
            });
        });
    try {
        await assert.rejects(
            scanner.loadSchedule({
                scheduleUrl: 'https://example.invalid/schedule.json',
                fetchTimeoutMs: 20,
            }),
            /timed out after 20ms/,
        );
    } finally {
        global.fetch = originalFetch;
    }
});

test('reports unresolved Azure parameter and macro selector references', () => {
    const repoRoot = makeTempRepo();
    writeFile(
        repoRoot,
        '.azuredevops/pipelines/external-param.yml',
        `
steps:
  - task: UseNode@1
    inputs:
      version: \${{ parameters.nodeVersion }}
`,
    );
    writeFile(
        repoRoot,
        '.azuredevops/pipelines/external-macro.yml',
        `
steps:
  - task: UseNode@1
    inputs:
      version: "$(nodeVersion)"
`,
    );

    const inventory = scanner.collectNodeSelectors(repoRoot);

    // Neither reference resolves from checked-in YAML, so each is reported as a
    // problem instead of being silently dropped, and no selector is invented.
    assert.equal(inventory.problems.length, 2);
    assert(
        inventory.problems.every((problem) =>
            /cannot be verified from checked-in YAML/.test(problem.message),
        ),
    );
    assert.equal(inventory.selectors.length, 0);
});

function writeNodeWorkflow(repoRoot, inputs, matrix) {
    writeFile(repoRoot, '.github/workflows/node.yml', yaml.stringify({
        jobs: { test: {
            ...(matrix ? { strategy: { matrix } } : {}),
            steps: [{ uses: 'actions/setup-node@v7', with: inputs }],
        } },
    }));
}

test('Azure Node tasks report missing selectors and preserve explicit siblings', () => {
    const cases = [
        ['UseNode@1', undefined, 'version'],
        ['UseNode@1', {}, 'version'],
        ['NodeTool@0', undefined, 'versionSpec'],
        ['NodeTool@0', {}, 'versionSpec'],
        ['NodeTool@0', { versionSource: 'fromFile' }, 'versionFilePath'],
        ['NodeTool@0', { versionSource: 'fromFile', versionSpec: '24' }, 'versionFilePath'],
    ];
    for (const [task, inputs, missing] of cases) {
        const repoRoot = makeTempRepo();
        writeFile(repoRoot, '.azuredevops/pipelines/node.yml', yaml.stringify({ steps: [
            { task, inputs },
            { task: 'UseNode@1', inputs: { version: '24' } },
        ] }));
        const inventory = scanner.collectNodeSelectors(repoRoot);
        assert.deepEqual(inventory.selectors.map((item) => item.rawValue), ['24']);
        assert.equal(inventory.problems.length, 1);
        assert(inventory.problems[0].message.includes(`no checked-in ${missing} input`));
    }
});

test('missing Azure selector oracle detects guard removal', () => {
    const repoRoot = makeTempRepo();
    writeFile(repoRoot, '.azuredevops/pipelines/node.yml', yaml.stringify({ steps: [
        { task: 'UseNode@1' },
    ] }));
    assert.equal(scanner.collectNodeSelectors(repoRoot).problems.length, 1);
    const original = fs.readFileSync(path.resolve(__dirname, '../../.github/scripts/check-toolchain-eol.js'), 'utf8');
    const guard = 'if (!Object.hasOwn(inputs, selectedInput))';
    assert.equal(original.split(guard).length - 1, 1);
    const mutant = path.join(repoRoot, 'scanner-mutant.js');
    fs.writeFileSync(mutant, original.replace(guard, 'if (false)'));
    const child = spawnSync(process.execPath, ['-e', `
        const assert = require('assert/strict');
        const inventory = require(process.argv[1]).collectNodeSelectors(process.argv[2]);
        assert.equal(inventory.problems.length, 1, 'An implicit Azure task default must not pass inventory.');
    `, mutant, repoRoot], {
        encoding: 'utf8', timeout: 30000,
        env: { ...process.env, NODE_PATH: path.resolve(__dirname, '../../node_modules') },
    });
    assert.ifError(child.error);
    assert.equal(child.status, 1);
    assert.match(child.stderr, /AssertionError/);
    assert.match(child.stderr, /An implicit Azure task default must not pass inventory/);
});

test('CLI returns native failure for missing GitHub and Azure Node selectors', () => {
    for (const host of ['github', 'azure']) {
        const repoRoot = makeTempRepo();
        if (host === 'github') {
            writeNodeWorkflow(repoRoot, {});
        } else {
            writeFile(repoRoot, '.azuredevops/pipelines/node.yml', yaml.stringify({ steps: [
                { task: 'NodeTool@0' },
            ] }));
        }
        const child = spawnSync(process.execPath, [
            path.resolve(__dirname, '../../.github/scripts/check-toolchain-eol.js'),
            '--repo-root', repoRoot, '--schedule-file',
            path.join(__dirname, 'fixtures/node-schedule.json'), '--json',
        ], { encoding: 'utf8', timeout: 30000 });
        assert.ifError(child.error);
        assert.equal(child.status, 1, child.stderr);
        const report = JSON.parse(child.stdout);
        assert.equal(report.problems.length, 1);
        assert.deepEqual(report.selectors, []);
    }
});

const nodeVersionFileCases = [
    ['.nvmrc', '24.18.0\n', '24.18.0'],
    ['.node-version', '24.18.0\n', '24.18.0'],
    ['.tool-versions', 'python 3.13.14\nnodejs 24.18.0\n', '24.18.0'],
    ['.tool-versions', 'node 24.18.0\n', '24.18.0'],
    ['.nvmrc', 'v24.18.0\n', '24.18.0'],
    ['config/package.json', { volta: { node: '22.19.0' }, devEngines: { runtime: { name: 'node', version: '24.18.0' } }, engines: { node: '26.9.0' } }, '22.19.0'],
    ['config/package.json', { devEngines: { runtime: { name: 'NoDe', version: '24.18.0' } }, engines: { node: '26.9.0' } }, '24.18.0'],
    ['config/package.json', { devEngines: { runtime: [{ name: 'python', version: '3.13.14' }, { name: 'node' }, { name: 'NODE', version: '22.19.0' }, { name: 'node', version: '24.18.0' }] }, engines: { node: '26.9.0' } }, '22.19.0'],
    ['config/package.json', { engines: { node: '24.18.0' }, volta: { extends: 'missing.json' } }, '24.18.0'],
    ['config/package.json', { engines: { node: '24' } }, '24'],
    ['config/package.json', { volta: { node: '24.18.0' }, devEngines: { runtime: [null] } }, '24.18.0'],
];

for (const [index, [file, value, expected]] of nodeVersionFileCases.entries()) {
    test(`inventories setup-node file selection ${index}: ${file}`, () => {
        const repoRoot = makeTempRepo();
        writeFile(repoRoot, file, typeof value === 'string' ? value : JSON.stringify(value));
        writeNodeWorkflow(repoRoot, { 'node-version-file': file });
        const inventory = scanner.collectNodeSelectors(repoRoot);
        assert.deepEqual(inventory.problems, []);
        assert.deepEqual(inventory.selectors.map((item) => [item.rawValue, item.referencedPath]), [[expected, file]]);
    });
}

test('setup-node inventory follows relative Volta inheritance', () => {
    const repoRoot = makeTempRepo();
    writeFile(repoRoot, 'config/package.json', JSON.stringify({ volta: { extends: '../shared/base.json' } }));
    writeFile(repoRoot, 'shared/base.json', JSON.stringify({ volta: { extends: 'runtime.json' } }));
    writeFile(repoRoot, 'shared/runtime.json', JSON.stringify({ engines: { node: '24.18.0' } }));
    writeNodeWorkflow(repoRoot, { 'node-version-file': 'config/package.json' });
    const inventory = scanner.collectNodeSelectors(repoRoot);
    assert.deepEqual(inventory.problems, []);
    assert.deepEqual(inventory.selectors.map((item) => [item.rawValue, item.referencedPath]), [['24.18.0', 'shared/runtime.json']]);
});

test('nonempty direct setup-node input overrides the file; blank direct input permits it', () => {
    for (const [direct, file, expected] of [
        ['24', 'missing.json', '24'],
        ['24', '.nvmrc', '24'],
        ['', '.nvmrc', '22.19.0'],
        ['   ', '.nvmrc', '22.19.0'],
    ]) {
        const repoRoot = makeTempRepo();
        writeFile(repoRoot, '.nvmrc', '22.19.0\n');
        writeNodeWorkflow(repoRoot, { 'node-version': direct, 'node-version-file': file });
        const inventory = scanner.collectNodeSelectors(repoRoot);
        assert.deepEqual(inventory.problems, []);
        assert.deepEqual(inventory.selectors.map((item) => item.rawValue), [expected]);
    }
});

test('setup-node file inventory reports missing, malformed, escaping, and cyclic sources', () => {
    const cases = [
        ['missing.json', undefined, /ENOENT/],
        ['config/package.json', '{', /valid JSON/],
        ['config/package.json', {}, /does not select a string/],
        ['config/package.json', { volta: { node: 24 } }, /does not select a string/],
        ['config/package.json', { devEngines: { runtime: [null] }, engines: { node: '24.18.0' } }, /string name/],
        ['config/package.json', { devEngines: { runtime: { name: 42, version: '24.18.0' } }, engines: { node: '24.18.0' } }, /string name/],
        ['config/package.json', { volta: { extends: 42 } }, /repository-contained/],
        ['config/package.json', { volta: { extends: '../../outside.json' } }, /escapes repository root/],
        ['config/package.json', { volta: { extends: 'package.json' } }, /cyclic/],
        ['.nvmrc', '', /does not select a string/],
    ];
    for (const [file, value, message] of cases) {
        const repoRoot = makeTempRepo();
        if (value !== undefined) {
            writeFile(repoRoot, file, typeof value === 'string' ? value : JSON.stringify(value));
        }
        writeNodeWorkflow(repoRoot, { 'node-version-file': file });
        const inventory = scanner.collectNodeSelectors(repoRoot);
        assert.equal(inventory.selectors.length, 0);
        assert.equal(inventory.problems.length, 1);
        assert.match(inventory.problems[0].message, message);
    }
});

test('setup-node mixed matrix variants retain the literal version-file fallback', () => {
    const cases = [
        [{ node: ['', '24'] }, ['22.19.0', '24']],
        [{ node: [null, '24'] }, ['22.19.0', '24']],
        [{ include: [{ node: ' ' }, { node: '24' }] }, ['22.19.0', '24']],
        [{ node: ['', ' '] }, ['22.19.0']],
        [{ node: ['24', '26'] }, ['24', '26']],
    ];
    for (const [matrix, expected] of cases) {
        const repoRoot = makeTempRepo();
        writeFile(repoRoot, '.nvmrc', '22.19.0\n');
        writeNodeWorkflow(repoRoot, {
            'node-version': '${{ matrix.node }}',
            'node-version-file': expected.includes('22.19.0') ? '.nvmrc' : 'ignored-missing-file',
        }, matrix);
        const inventory = scanner.collectNodeSelectors(repoRoot);
        assert.deepEqual(inventory.problems, []);
        assert.deepEqual(inventory.selectors.map((item) => item.rawValue).sort(), expected);
    }
});

test('setup-node all-blank direct matrix still inventories file-matrix values', () => {
    const repoRoot = makeTempRepo();
    writeFile(repoRoot, '.nvmrc', '22.19.0\n');
    writeFile(repoRoot, '.node-version', '24.18.0\n');
    writeNodeWorkflow(repoRoot, {
        'node-version': '${{ matrix.node }}',
        'node-version-file': '${{ matrix.file }}',
    }, { node: ['', ' '], file: ['.nvmrc', '.node-version'] });
    const inventory = scanner.collectNodeSelectors(repoRoot);
    assert.deepEqual(inventory.problems, []);
    assert.deepEqual(inventory.selectors.map((item) => item.rawValue).sort(), ['22.19.0', '24.18.0']);
});

test('setup-node correlates static pairs and treats missing direct properties as blank', () => {
    for (const [matrix, expected] of [
        [{ include: [{ node: '', file: '.nvmrc' }, { node: '24', file: 'ignored-file' }] }, ['22.19.0', '24']],
        [{ file: ['.nvmrc'] }, ['22.19.0']],
    ]) {
        const repoRoot = makeTempRepo();
        writeFile(repoRoot, '.nvmrc', '22.19.0\n');
        writeNodeWorkflow(repoRoot, {
            'node-version': '${{ matrix.node }}',
            'node-version-file': '${{ matrix.file }}',
        }, matrix);
        const inventory = scanner.collectNodeSelectors(repoRoot);
        assert.deepEqual(inventory.selectors.map((item) => item.rawValue), expected);
        assert.deepEqual(inventory.problems, []);
    }
});

test('setup-node reports unresolved consulted version-file matrices', () => {
    for (const matrix of [{}, { file: [] }]) {
        for (const direct of [undefined, '${{ matrix.node }}']) {
            const repoRoot = makeTempRepo();
            writeNodeWorkflow(repoRoot, {
                ...(direct === undefined ? {} : { 'node-version': direct }),
                'node-version-file': '${{ matrix.file }}',
            }, { ...matrix, node: ['', ' '] });
            const inventory = scanner.collectNodeSelectors(repoRoot);
            assert.deepEqual(inventory.selectors, []);
            assert.equal(inventory.problems.length, 1);
            assert.equal(inventory.problems[0].path.split(path.sep).join('/'), '.github/workflows/node.yml');
            assert.match(inventory.problems[0].message, Object.hasOwn(matrix, 'file')
                ? /axis file must be a nonempty static list/ : /version-file input has a blank value/);
        }
    }
});

test('setup-node reports blank consulted version-file values without losing valid files', () => {
    for (const blank of ['', ' ', null]) {
        for (const validFile of [false, true]) {
            const repoRoot = makeTempRepo();
            writeFile(repoRoot, '.nvmrc', '22.19.0\n');
            writeNodeWorkflow(repoRoot, { 'node-version-file': '${{ matrix.file }}' }, {
                file: validFile ? [blank, '.nvmrc'] : [blank],
            });
            const inventory = scanner.collectNodeSelectors(repoRoot);
            assert.deepEqual(inventory.selectors.map((item) => item.rawValue), validFile ? ['22.19.0'] : []);
            assert.equal(inventory.problems.length, 1);
            assert.match(inventory.problems[0].message, /version-file input has a blank value/);
        }
        const repoRoot = makeTempRepo();
        writeNodeWorkflow(repoRoot, { 'node-version-file': blank });
        const inventory = scanner.collectNodeSelectors(repoRoot);
        assert.deepEqual(inventory.selectors, []);
        assert.equal(inventory.problems.length, 1);
        assert.match(inventory.problems[0].message, /version-file input has a blank value/);
    }
});

test('setup-node nonblank direct input ignores unresolved or blank version-file inputs', () => {
    for (const file of ['${{ matrix.file }}', '', null]) {
        const repoRoot = makeTempRepo();
        writeNodeWorkflow(repoRoot, {
            'node-version': '${{ matrix.node }}',
            'node-version-file': file,
        }, { node: ['24'] });
        const inventory = scanner.collectNodeSelectors(repoRoot);
        assert.deepEqual(inventory.problems, []);
        assert.deepEqual(inventory.selectors.map((item) => item.rawValue), ['24']);
    }
});

test('setup-node excludes unreachable blank variants before reading the version file', () => {
    for (const file of ['.nvmrc', 'missing-file']) {
        const repoRoot = makeTempRepo();
        writeFile(repoRoot, '.nvmrc', '16.0.0\n');
        writeNodeWorkflow(repoRoot, {
            'node-version': '${{ matrix.node }}', 'node-version-file': file,
        }, { node: ['', '24'], exclude: [{ node: '' }] });
        const inventory = scanner.collectNodeSelectors(repoRoot);
        assert.deepEqual(inventory.problems, []);
        assert.deepEqual(inventory.selectors.map((item) => item.rawValue), ['24']);
    }
});

test('setup-node distinguishes partial exclusions, complete removal, and include restoration', () => {
    const cases = [
        [{ os: ['ubuntu', 'windows'], node: ['', '24'], exclude: [{ os: 'ubuntu', node: '' }] }, ['22.19.0', '24']],
        [{ os: ['ubuntu', 'windows'], node: ['', '24'], exclude: [{ node: '' }] }, ['24']],
        [{ node: ['', '24'], exclude: [{ node: '' }], include: [{ node: '' }] }, ['22.19.0', '24']],
        [{ node: [''], exclude: [{ node: '' }] }, []],
        [{ node: [''], exclude: [{ node: '' }], include: [{ node: '24' }] }, ['24']],
    ];
    for (const [matrix, expected] of cases) {
        const repoRoot = makeTempRepo();
        if (expected.includes('22.19.0')) writeFile(repoRoot, '.nvmrc', '22.19.0\n');
        writeNodeWorkflow(repoRoot, {
            'node-version': '${{ matrix.node }}', 'node-version-file': '.nvmrc',
        }, matrix);
        const inventory = scanner.collectNodeSelectors(repoRoot);
        assert.deepEqual(inventory.problems, []);
        assert.deepEqual(inventory.selectors.map((item) => item.rawValue).sort(), expected);
    }
});

test('setup-node applies ordered includes to original combinations only', () => {
    const repoRoot = makeTempRepo();
    writeFile(repoRoot, '.node-version', '22.19.0\n');
    writeFile(repoRoot, '.nvmrc', '24.18.0\n');
    writeNodeWorkflow(repoRoot, {
        'node-version': '${{ matrix.node }}', 'node-version-file': '${{ matrix.file }}',
    }, {
        os: ['ubuntu'], node: [''],
        include: [
            { file: 'overwritten-missing-file' },
            { os: 'windows', node: '', file: '.nvmrc' },
            { file: '.node-version' },
        ],
    });
    const inventory = scanner.collectNodeSelectors(repoRoot);
    assert.deepEqual(inventory.problems, []);
    assert.deepEqual(inventory.selectors.map((item) => [item.rawValue, item.referencedPath]), [
        ['22.19.0', '.node-version'], ['24.18.0', '.nvmrc'],
    ]);
});

test('setup-node resolves nested object axes and compares excluded objects structurally', () => {
    const repoRoot = makeTempRepo();
    writeNodeWorkflow(repoRoot, {
        'node-version': '${{ matrix.node.version }}', 'node-version-file': 'missing-file',
    }, {
        node: [{ version: '', env: 'test' }, { version: '24' }],
        exclude: [{ node: { env: 'test', version: '' } }],
    });
    const inventory = scanner.collectNodeSelectors(repoRoot);
    assert.deepEqual(inventory.problems, []);
    assert.deepEqual(inventory.selectors.map((item) => item.rawValue), ['24']);
});

test('setup-node reports missing direct matrix properties without a file fallback', () => {
    const cases = [
        [undefined, '${{ matrix.node }}', []],
        [{ os: ['ubuntu'] }, '${{ matrix.node }}', []],
        [{ node: [{}] }, '${{ matrix.node.version }}', []],
        [{ node: [null] }, '${{ matrix.node.version }}', []],
        [{ include: [{ os: 'ubuntu' }, { os: 'windows', node: '24' }] }, '${{ matrix.node }}', ['24']],
        [{ os: ['ubuntu', 'windows'] }, '${{ matrix.node }}', []],
    ];
    for (const [matrix, direct, expected] of cases) {
        const repoRoot = makeTempRepo();
        writeNodeWorkflow(repoRoot, { 'node-version': direct }, matrix);
        const inventory = scanner.collectNodeSelectors(repoRoot);
        assert.deepEqual(inventory.selectors.map((item) => item.rawValue), expected);
        assert.equal(inventory.problems.length, 1);
        assert.match(inventory.problems[0].message, /direct matrix input matrix\.node(?:\.version)? is missing.*no version-file fallback/);
        assert.equal(inventory.problems[0].path.split(path.sep).join('/'), '.github/workflows/node.yml');
    }
});

test('setup-node preserves missing-property fallback and excluded combinations', () => {
    for (const file of ['.nvmrc', '${{ matrix.file }}']) {
        const repoRoot = makeTempRepo();
        writeFile(repoRoot, '.nvmrc', '22.19.0\n');
        writeNodeWorkflow(repoRoot, { 'node-version': '${{ matrix.node }}', 'node-version-file': file },
            { file: ['.nvmrc'] });
        const inventory = scanner.collectNodeSelectors(repoRoot);
        assert.deepEqual(inventory.problems, []);
        assert.deepEqual(inventory.selectors.map((item) => item.rawValue), ['22.19.0']);
    }
    const repoRoot = makeTempRepo();
    writeNodeWorkflow(repoRoot, { 'node-version': '${{ matrix.node }}' },
        { os: ['ubuntu'], exclude: [{ os: 'ubuntu' }] });
    assert.deepEqual(scanner.collectNodeSelectors(repoRoot), { selectors: [], problems: [] });
});

test('setup-node reports blank and auth-only inputs without losing valid siblings', () => {
    for (const empty of ['', ' ', null]) {
        const repoRoot = makeTempRepo();
        writeNodeWorkflow(repoRoot, { 'node-version': '${{ matrix.node }}' }, { node: [empty, '24'] });
        const inventory = scanner.collectNodeSelectors(repoRoot);
        assert.equal(inventory.problems.length, 1);
        assert.match(inventory.problems[0].message, /no nonblank node-version or version-file fallback/);
        assert.deepEqual(inventory.selectors.map((item) => item.rawValue), ['24']);
        writeNodeWorkflow(repoRoot, { 'node-version': empty });
        const blank = scanner.collectNodeSelectors(repoRoot);
        assert.deepEqual(blank.selectors, []);
        assert.equal(blank.problems.length, 1);
        assert.match(blank.problems[0].message, /no checked-in runtime can be inventoried/);
    }
    for (const inputs of [undefined, {}, { 'registry-url': 'https://registry.npmjs.org' }]) {
        const repoRoot = makeTempRepo();
        writeNodeWorkflow(repoRoot, inputs);
        const inventory = scanner.collectNodeSelectors(repoRoot);
        assert.deepEqual(inventory.selectors, []);
        assert.equal(inventory.problems.length, 1);
        assert.match(inventory.problems[0].message, /no nonblank node-version or version-file fallback/);
        assert.equal(inventory.problems[0].path.split(path.sep).join('/'), '.github/workflows/node.yml');
    }
});

test('missing literal selector oracle detects diagnostic removal', () => {
    const repoRoot = makeTempRepo();
    writeNodeWorkflow(repoRoot, {});
    assert.equal(scanner.collectNodeSelectors(repoRoot).problems.length, 1);
    const original = fs.readFileSync(path.resolve(__dirname, '../../.github/scripts/check-toolchain-eol.js'), 'utf8');
    const diagnostic = "throw new Error('Node.js setup input has no nonblank node-version or version-file fallback; no checked-in runtime can be inventoried.');";
    assert.equal(original.split(diagnostic).length - 1, 1);
    const mutant = path.join(repoRoot, 'scanner-mutant.js');
    fs.writeFileSync(mutant, original.replace(diagnostic, 'continue;'));
    const child = spawnSync(process.execPath, ['-e', `
        const assert = require('assert/strict');
        const inventory = require(process.argv[1]).collectNodeSelectors(process.argv[2]);
        assert.equal(inventory.problems.length, 1, 'An implicit PATH runtime must not pass inventory.');
    `, mutant, repoRoot], {
        encoding: 'utf8', timeout: 30000,
        env: { ...process.env, NODE_PATH: path.resolve(__dirname, '../../node_modules') },
    });
    assert.ifError(child.error);
    assert.equal(child.status, 1);
    assert.match(child.stderr, /AssertionError/);
    assert.match(child.stderr, /An implicit PATH runtime must not pass inventory/);
});

test('missing matrix diagnostic oracle detects assertion removal', () => {
    const repoRoot = makeTempRepo();
    writeNodeWorkflow(repoRoot, { 'node-version': '${{ matrix.node }}' },
        { include: [{ os: 'ubuntu' }, { os: 'windows', node: '24' }] });
    const baseline = scanner.collectNodeSelectors(repoRoot);
    assert.deepEqual(baseline.selectors.map((item) => item.rawValue), ['24']);
    assert.equal(baseline.problems.length, 1);
    const original = fs.readFileSync(path.resolve(__dirname, '../../.github/scripts/check-toolchain-eol.js'), 'utf8');
    const predicate = 'if (resolved.missingProperty)';
    assert.equal(original.split(predicate).length - 1, 1);
    const mutant = path.join(repoRoot, 'scanner-mutant.js');
    fs.writeFileSync(mutant, original.replace(predicate, 'if (false)'));
    const child = spawnSync(process.execPath, ['-e', `
        const assert = require('assert/strict');
        const inventory = require(process.argv[1]).collectNodeSelectors(process.argv[2]);
        assert.deepEqual(inventory.selectors.map((item) => item.rawValue), ['24']);
        assert.equal(inventory.problems.length, 1, 'A valid sibling must not hide missing selection.');
        assert.match(inventory.problems[0].message, /direct matrix input matrix.node is missing/,
            'Missing matrix properties need their specific diagnostic.');
    `, mutant, repoRoot], {
        encoding: 'utf8', timeout: 30000,
        env: { ...process.env, NODE_PATH: path.resolve(__dirname, '../../node_modules') },
    });
    assert.ifError(child.error);
    assert.equal(child.status, 1);
    assert.match(child.stderr, /AssertionError/);
    assert.match(child.stderr, /Missing matrix properties need their specific diagnostic/);
});

test('setup-node static matrix diagnostics retain unknown and malformed inputs', () => {
    const cases = [
        ['${{ fromJSON(needs.build.outputs.matrix) }}', /static matrix mapping/],
        [{ node: '${{ needs.build.outputs.versions }}' }, /dynamic/],
        [{ node: ['${{ vars.NODE }}'] }, /dynamic/],
        [{ node: ['24'], include: {} }, /include must be a list/],
        [{ node: ['24'], exclude: [null] }, /exclude must be a list/],
        [{ node: [{ version: '24' }] }, /must be scalar/],
        [{ node: ['24'], os: [] }, /nonempty static list/],
    ];
    for (const [matrix, message] of cases) {
        const repoRoot = makeTempRepo();
        writeNodeWorkflow(repoRoot, { 'node-version': '${{ matrix.node }}' }, matrix);
        const inventory = scanner.collectNodeSelectors(repoRoot);
        assert.deepEqual(inventory.selectors, []);
        assert.equal(inventory.problems.length, 1);
        assert.match(inventory.problems[0].message, message);
        assert.equal(inventory.problems[0].path.split(path.sep).join('/'), '.github/workflows/node.yml');
    }
});

test('setup-node matrix limits apply separately before and after exclusions', () => {
    const cases = [
        [{ node: ['24'], job: Array.from({ length: 256 }, (_, i) => i) }, null],
        [{ node: ['24'], job: Array.from({ length: 257 }, (_, i) => i) }, /256-job/],
        [{ node: ['24'], os: ['keep', 'remove'], job: Array.from({ length: 256 }, (_, i) => i), exclude: [{ os: 'remove' }] }, null],
        [{ node: ['24'], job: Array.from({ length: 4097 }, (_, i) => i), exclude: [{}] }, /4096-combination/],
        [{ include: Array.from({ length: 257 }, (_, i) => ({ node: '24', job: i })) }, /256-job/],
    ];
    for (const [matrix, message] of cases) {
        const repoRoot = makeTempRepo();
        writeNodeWorkflow(repoRoot, { 'node-version': '${{ matrix.node }}' }, matrix);
        const inventory = scanner.collectNodeSelectors(repoRoot);
        if (message) {
            assert.deepEqual(inventory.selectors, []);
            assert.equal(inventory.problems.length, 1);
            assert.match(inventory.problems[0].message, message);
        } else {
            assert.deepEqual(inventory.problems, []);
            assert.deepEqual(inventory.selectors.map((item) => item.rawValue), ['24']);
        }
    }
});

test('setup-node literal direct input ignores an unrelated dynamic matrix', () => {
    const repoRoot = makeTempRepo();
    writeNodeWorkflow(repoRoot, { 'node-version': '24', 'node-version-file': 'missing-file' },
        '${{ fromJSON(needs.build.outputs.matrix) }}');
    const inventory = scanner.collectNodeSelectors(repoRoot);
    assert.deepEqual(inventory.problems, []);
    assert.deepEqual(inventory.selectors.map((item) => item.rawValue), ['24']);
});

test('setup-node matrix regression oracles fail when exclusion or correlation is removed', () => {
    const originalPath = path.resolve(__dirname, '../../.github/scripts/check-toolchain-eol.js');
    const original = fs.readFileSync(originalPath, 'utf8');
    const cases = [
        [
            /originals = originals\.filter\([\s\S]*?matrixValuesEqual\(entry\[key\], value\)\)\)\);/,
            'originals = originals;',
            { node: ['', '24'], exclude: [{ node: '' }] }, ['24'],
        ],
        [
            /resolveGithubExpression\(file, combination\)/,
            'resolveGithubExpression(file, combinations[0])',
            { include: [{ node: '24', file: 'ignored-file' }, { node: '', file: '.nvmrc' }] }, ['24', '22.19.0'],
        ],
    ];
    for (const [pattern, replacement, matrix, expected] of cases) {
        const repoRoot = makeTempRepo();
        writeFile(repoRoot, '.nvmrc', '22.19.0\n');
        writeNodeWorkflow(repoRoot, {
            'node-version': '${{ matrix.node }}',
            'node-version-file': Object.hasOwn(matrix, 'include') ? '${{ matrix.file }}' : '.nvmrc',
        }, matrix);
        const inventory = scanner.collectNodeSelectors(repoRoot);
        assert.deepEqual(inventory.problems, []);
        assert.deepEqual(inventory.selectors.map((item) => item.rawValue), expected);
        assert.match(original, pattern);
        const mutant = path.join(repoRoot, 'scanner-mutant.js');
        fs.writeFileSync(mutant, original.replace(pattern, replacement));
        const child = spawnSync(process.execPath, ['-e', `
            const assert = require('assert/strict');
            const inventory = require(process.argv[1]).collectNodeSelectors(process.argv[2]);
            assert.deepEqual(inventory.problems, []);
            assert.deepEqual(inventory.selectors.map((item) => item.rawValue), JSON.parse(process.argv[3]));
        `, mutant, repoRoot, JSON.stringify(expected)], {
            encoding: 'utf8', timeout: 30000,
            env: { ...process.env, NODE_PATH: path.resolve(__dirname, '../../node_modules') },
        });
        assert.equal(child.error, undefined);
        assert.equal(child.status, 1);
        assert.match(child.stderr, /AssertionError/);
    }
});

test('setup-node inventory accepts 32 inherited files and reports a longer chain', () => {
    for (const count of [32, 33]) {
        const repoRoot = makeTempRepo();
        for (let index = 0; index < count; index += 1) {
            writeFile(repoRoot, `config/${index}.json`, JSON.stringify(index + 1 === count
                ? { engines: { node: '24.18.0' } }
                : { volta: { extends: `${index + 1}.json` } }));
        }
        writeNodeWorkflow(repoRoot, { 'node-version-file': 'config/0.json' });
        const inventory = scanner.collectNodeSelectors(repoRoot);
        if (count === 32) {
            assert.deepEqual(inventory.problems, []);
            assert.deepEqual(inventory.selectors.map((item) => item.rawValue), ['24.18.0']);
        } else {
            assert.equal(inventory.selectors.length, 0);
            assert.equal(inventory.problems.length, 1);
            assert.match(inventory.problems[0].message, /32-file/);
        }
    }
});

const yamlGuidePath = path.resolve(__dirname, '..', '..', '.github/instructions/yaml.instructions.md');
const optionalYamlGuide = { skip: !fs.existsSync(yamlGuidePath) && 'Optional YAML guide is excluded.' };

function documentedNodeExample() {
    const guide = fs.readFileSync(yamlGuidePath, 'utf8');
    const section = guide.split(/### Exact Node\.js version-file exception\r?\n/)[1]
        .split(/\r?\n## GitHub Actions Documentation Comment URLs/)[0];
    const manifest = JSON.parse(section.match(/```json\r?\n([\s\S]*?)\r?\n```/)[1]);
    const steps = yaml.parse(section.match(/```yaml\r?\n([\s\S]*?)\r?\n```/)[1]);
    return { manifest, steps };
}

function assertDocumentedNodeExample(manifest, steps) {
    // A narrow oracle for this documented example, not an arbitrary-workflow validator.
    assert.deepEqual(Object.keys(manifest), ['engines']);
    assert.match(manifest.engines.node, /^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$/);
    const setup = steps.findIndex((step) => step.uses && step.uses.startsWith('actions/setup-node@'));
    const verify = steps.findIndex((step) => step.name === 'Verify the exact Node.js runtime');
    const install = steps.findIndex((step) => step.run === 'npm ci');
    assert(setup >= 0 && verify > setup && install > verify);
    assert.equal(steps[setup].with['node-version-file'], 'package.json');
    assert(!Object.hasOwn(steps[setup].with, 'node-version'));
    assert.equal(steps[setup].with['package-manager-cache'], false);
    assert.equal(steps[verify].shell, 'bash');
    assert.match(steps[verify].run, /readFileSync\('package\.json', 'utf8'\)\)\.engines\.node/);
    assert.match(steps[verify].run, /process\.versions\.node !== expected/);
}

test('documented Node file example preserves selection and verification order', optionalYamlGuide, () => {
    const { manifest, steps } = documentedNodeExample();
    assertDocumentedNodeExample(manifest, steps);
    const mutations = [
        (m) => { m.engines.node = '24'; },
        (m) => { m.engines.node = '24.x'; },
        (m) => { m.engines.node = '>=24'; },
        (m) => { m.engines.node = 'lts/*'; },
        (m) => { m.engines.node = '24.18.0-rc.1'; },
        (m) => { m.engines.node = '24.18.0+build'; },
        (m, s) => { s[1].with['node-version'] = '24'; },
        (m, s) => { s[2].run = s[2].run.replace('.engines.node', '.volta.node'); },
        (m, s) => { s[2].run = 'node --version'; },
        (m, s) => { [s[2], s[3]] = [s[3], s[2]]; },
    ];
    for (const mutate of mutations) {
        const changed = structuredClone({ manifest, steps });
        mutate(changed.manifest, changed.steps);
        assert.throws(() => assertDocumentedNodeExample(changed.manifest, changed.steps));
    }
});

test('documented Node verifier executes and rejects mismatched or non-exact values', optionalYamlGuide, () => {
    const { steps } = documentedNodeExample();
    const verification = steps.find((step) => step.name === 'Verify the exact Node.js runtime');
    const script = verification.run.match(/^node <<'NODE'\n([\s\S]*)\nNODE\n?$/)[1];
    const repoRoot = makeTempRepo();
    // Use the executing runtime for a portable success; the published pin is illustrative.
    const cases = [
        [process.versions.node, 0],
        ['0.0.0', 1],
        ['24', 1],
        ['24.x', 1],
        ['>=24', 1],
        ['lts/*', 1],
        ['24.18.0-rc.1', 1],
        ['24.18.0+build', 1],
    ];
    for (const [expected, exitCode] of cases) {
        writeFile(repoRoot, 'package.json', JSON.stringify({ engines: { node: expected } }));
        const result = spawnSync(process.execPath, ['-e', script], { cwd: repoRoot, encoding: 'utf8', timeout: 10000 });
        assert.ifError(result.error);
        assert.equal(result.status, exitCode, result.stderr);
        if (exitCode !== 0) {
            assert.match(result.stderr, /Expected Node.js/);
        }
    }
});


function azureNode(version, task = 'UseNode@1') {
    return { steps: [{ task, inputs: { [task === 'UseNode@1' ? 'version' : 'versionSpec']: version } }] };
}

function writeAzure(repoRoot, filename, value) {
    writeFile(repoRoot, filename, yaml.stringify(value));
}

function azureCli(repoRoot, paths = [], executable) {
    const schedule = path.join(repoRoot, 'fixed-schedule.json');
    fs.writeFileSync(schedule, JSON.stringify({ v18: { end: '2025-04-30' }, v24: { end: '2030-04-30' } }));
    const child = spawnSync(process.execPath, [
        executable || path.resolve(__dirname, '../../.github/scripts/check-toolchain-eol.js'),
        '--repo-root', repoRoot, '--schedule-file', schedule, '--as-of', '2026-09-22', '--json',
        ...paths.flatMap((filename) => ['--azure-pipeline', filename]),
    ], { cwd: repoRoot, encoding: 'utf8', timeout: 30000,
        env: { ...process.env, NODE_PATH: path.resolve(__dirname, '../../node_modules') } });
    assert.ifError(child.error);
    assert.equal(child.stderr, '');
    return { status: child.status, report: JSON.parse(child.stdout) };
}

function parameterizedAzure(defaultVersion) {
    return { parameters: [{ name: 'node', type: 'string', default: defaultVersion }],
        ...azureNode('${{ parameters.node }}') };
}

for (const [name, caller, target] of [
    ['extends', { extends: { template: 'ci/build.yml' } }, azureNode('18')],
    ['steps', { steps: [{ template: 'ci/build.yml' }] }, azureNode('18')],
    ['jobs', { jobs: [{ template: 'ci/build.yml' }] }, { jobs: [{ job: 'build', ...azureNode('18', 'NodeTool@0') }] }],
    ['stages', { stages: [{ template: 'ci/build.yml' }] }, { stages: [{ stage: 'build', jobs: [{ job: 'build', ...azureNode('18') }] }] }],
]) {
    test(`Azure local ${name} templates retain actual path and EOL native failure`, () => {
        const repoRoot = makeTempRepo();
        writeAzure(repoRoot, 'azure-pipelines.yml', caller);
        writeAzure(repoRoot, 'ci/build.yml', target);
        const result = azureCli(repoRoot);
        assert.equal(result.status, 1);
        assert.deepEqual(result.report.problems, []);
        assert.deepEqual(result.report.findings.map((item) => [item.path, item.rawValue, item.status]), [['ci/build.yml', '18', 'eol']]);
        writeFile(repoRoot, 'ci/build.yml', yaml.stringify(target).replace(/18/g, '24'));
        const positive = azureCli(repoRoot);
        assert.equal(positive.status, 0);
        assert.deepEqual(positive.report.problems, []);
        assert.deepEqual(positive.report.findings.map((item) => [item.path, item.rawValue, item.status]), [['ci/build.yml', '24', 'supported']]);
    });
}

for (const template of ['../shared/node.yml', '/shared/node.yml', '../shared/node.yml@self']) {
    test(`Azure template path resolves from the including file: ${template}`, () => {
        const repoRoot = makeTempRepo();
        writeAzure(repoRoot, 'azure-pipelines.yml', { steps: [{ template: 'ci/build.yml' }] });
        writeAzure(repoRoot, 'ci/build.yml', { steps: [{ template }] });
        writeAzure(repoRoot, 'shared/node.yml', azureNode('24'));
        const inventory = scanner.collectNodeSelectors(repoRoot);
        assert.deepEqual(inventory.problems, []);
        assert.deepEqual(inventory.selectors.map((item) => [item.path, item.rawValue]), [['shared/node.yml', '24']]);
    });
}

test('Azure implicit template defaults are not separate entrypoint invocations', () => {
    for (const [defaultVersion, overrideVersion] of [['24', '18'], ['18', '24']]) {
        const repoRoot = makeTempRepo();
        writeAzure(repoRoot, 'azure-pipelines.yml', { extends: {
            template: '.azuredevops/pipelines/build.yml', parameters: { node: overrideVersion },
        } });
        writeAzure(repoRoot, '.azuredevops/pipelines/build.yml', parameterizedAzure(defaultVersion));
        const result = azureCli(repoRoot);
        assert.equal(result.status, overrideVersion === '18' ? 1 : 0);
        assert.deepEqual(result.report.problems, []);
        assert.deepEqual(result.report.selectors.map((item) => item.rawValue), [overrideVersion]);
        const dual = azureCli(repoRoot, ['.azuredevops/pipelines/build.yml']);
        assert.equal(dual.status, 1);
        assert.deepEqual(dual.report.problems, []);
        assert.deepEqual(dual.report.selectors.map((item) => item.rawValue).sort(), ['18', '24']);
    }
});

test('Azure conventional root names remain independent roots when also called', () => {
    const repoRoot = makeTempRepo();
    writeAzure(repoRoot, 'azure-pipelines.yml', parameterizedAzure('18'));
    writeAzure(repoRoot, 'azure-pipelines.yaml', { extends: { template: 'azure-pipelines.yml', parameters: { node: '24' } } });
    const result = azureCli(repoRoot);
    assert.equal(result.status, 1);
    assert.deepEqual(result.report.problems, []);
    assert.deepEqual(result.report.selectors.map((item) => item.rawValue).sort(), ['18', '24']);
});

test('Azure repeated template calls and nested forwarding preserve distinct bindings', () => {
    for (const versions of [['18', '24'], ['24', '18']]) {
        const repoRoot = makeTempRepo();
        writeAzure(repoRoot, 'azure-pipelines.yml', { steps: versions.map((node) => ({ template: 'ci/outer.yml', parameters: { node } })) });
        writeAzure(repoRoot, 'ci/outer.yml', { parameters: [{ name: 'node', type: 'string' }],
            steps: [{ template: 'inner.yml', parameters: { node: '${{ parameters.node }}' } }] });
        writeAzure(repoRoot, 'ci/inner.yml', parameterizedAzure('24'));
        const result = azureCli(repoRoot);
        assert.equal(result.status, 1);
        assert.deepEqual(result.report.problems, []);
        assert.deepEqual(result.report.selectors.map((item) => [item.path, item.rawValue]), versions.map((node) => ['ci/inner.yml', node]));
    }
});

test('Azure variable templates resolve parameter scope and keep sibling overrides isolated', () => {
    const repoRoot = makeTempRepo();
    writeAzure(repoRoot, 'azure-pipelines.yml', { variables: [{ template: 'ci/vars.yml', parameters: { chosen: '24' } }],
        jobs: [{ job: 'old', variables: { node: '18' }, ...azureNode('$(node)') },
            { job: 'new', ...azureNode('$(node)') }] });
    writeAzure(repoRoot, 'ci/vars.yml', { parameters: [{ name: 'chosen', type: 'string', default: '18' }],
        variables: { node: '${{   parameters.chosen }}' } });
    const result = azureCli(repoRoot);
    assert.equal(result.status, 1);
    assert.deepEqual(result.report.problems, []);
    assert.deepEqual(result.report.selectors.map((item) => item.rawValue), ['18', '24']);
});

test('Azure template NodeTool version files retain workspace-relative paths', () => {
    const repoRoot = makeTempRepo();
    writeAzure(repoRoot, 'azure-pipelines.yml', { steps: [{ template: 'ci/node.yml' }] });
    writeAzure(repoRoot, 'ci/node.yml', { steps: [{ task: 'NodeTool@0', inputs: { versionSource: 'fromFile', versionFilePath: '.nvmrc' } }] });
    writeFile(repoRoot, '.nvmrc', '24\n');
    writeFile(repoRoot, 'ci/.nvmrc', '18\n');
    const inventory = scanner.collectNodeSelectors(repoRoot);
    assert.deepEqual(inventory.problems, []);
    assert.deepEqual(inventory.selectors.map((item) => [item.path, item.referencedPath, item.rawValue]), [['ci/node.yml', '.nvmrc', '24']]);
});

test('Azure custom entrypoints are additive, repeatable and explicit', () => {
    const repoRoot = makeTempRepo();
    writeAzure(repoRoot, 'azure-pipelines.yml', azureNode('24'));
    writeAzure(repoRoot, 'ci/one.yml', azureNode('18'));
    writeAzure(repoRoot, 'ci/two.yml', azureNode('24'));
    writeAzure(repoRoot, 'examples/unselected.yml', azureNode('18'));
    assert.deepEqual(scanner.collectNodeSelectors(repoRoot).selectors.map((item) => item.path), ['azure-pipelines.yml']);
    const result = azureCli(repoRoot, ['ci/one.yml', 'ci/two.yml', 'ci/one.yml']);
    assert.equal(result.status, 1);
    assert.deepEqual(result.report.problems, []);
    assert.deepEqual(result.report.selectors.map((item) => [item.path, item.rawValue]), [
        ['azure-pipelines.yml', '24'], ['ci/one.yml', '18'], ['ci/two.yml', '24'],
    ]);
    for (const args of [['--azure-pipeline'], ['--azure-pipeline', ''], ['--azure-pipeline', '--json']]) {
        assert.throws(() => scanner.parseArgs(args), /requires a repository-relative YAML path/);
    }
    const missing = azureCli(repoRoot, ['ci/missing.yml']);
    assert.equal(missing.status, 1);
    assert.match(missing.report.problems[0].message, /ENOENT/);
});

for (const [name, reference, body, diagnostic] of [
    ['missing', 'ci/missing.yml', null, /ENOENT/],
    ['malformed', 'ci/node.yml', 'steps: [\n', /Invalid Azure YAML/],
    ['remote', 'node.yml@elsewhere', null, /external template/],
    ['dynamic', '${{ parameters.filename }}', null, /literal local YAML path/],
    ['escaping', '../outside.yml', null, /escapes repository root/],
    ['role', 'ci/node.yml', 'jobs: []\n', /steps template has no steps/],
    ['conditional', 'ci/node.yml', 'steps:\n- ${{ if true }}:\n  - task: UseNode@1\n    inputs: {version: "18"}\n', /conditional or structural/],
    ['structural', 'ci/node.yml', 'steps: ${{ parameters.steps }}\n', /static list/],
]) {
    test(`Azure unverifiable ${name} references fail without hiding valid siblings`, () => {
        const repoRoot = makeTempRepo();
        writeAzure(repoRoot, 'azure-pipelines.yml', { steps: [{ template: reference }, ...azureNode('24').steps] });
        if (body !== null) writeFile(repoRoot, reference, body);
        const result = azureCli(repoRoot);
        assert.equal(result.status, 1);
        assert(result.report.problems.some((item) => diagnostic.test(item.message)));
        assert.deepEqual(result.report.selectors.map((item) => item.rawValue), ['24']);
    });
}

test('Azure rootless template cycles remain failures after implicit-root filtering', () => {
    const repoRoot = makeTempRepo();
    writeAzure(repoRoot, '.azuredevops/pipelines/a.yml', { extends: { template: 'b.yml' } });
    writeAzure(repoRoot, '.azuredevops/pipelines/b.yml', { extends: { template: 'a.yml' } });
    const result = azureCli(repoRoot);
    assert.equal(result.status, 1);
    assert(result.report.problems.some((item) => /cycle/.test(item.message)));
    assert.deepEqual(result.report.selectors, []);
});

test('Azure task input and script template text never causes file reads', () => {
    const repoRoot = makeTempRepo();
    writeAzure(repoRoot, 'azure-pipelines.yml', { steps: [
        { script: 'template: missing.yml' }, { task: 'Other@1', inputs: { template: 'missing.yml' } }, ...azureNode('24').steps,
    ] });
    const result = azureCli(repoRoot);
    assert.equal(result.status, 0);
    assert.deepEqual(result.report.problems, []);
    assert.deepEqual(result.report.selectors.map((item) => item.rawValue), ['24']);
});

test('Azure required, dynamic, cyclic and structural bindings do not fall back to defaults', () => {
    for (const argument of [undefined, '${{ parameters.absent }}', '$(absent)', { version: '24' }, ['24']]) {
        const repoRoot = makeTempRepo();
        writeAzure(repoRoot, 'azure-pipelines.yml', { extends: { template: 'ci/node.yml',
            parameters: argument === undefined ? {} : { node: argument } } });
        writeAzure(repoRoot, 'ci/node.yml', { parameters: [{ name: 'node', type: 'string' }], ...azureNode('${{ parameters.node }}') });
        const result = azureCli(repoRoot);
        assert.equal(result.status, 1);
        assert(result.report.problems.length > 0);
        assert.deepEqual(result.report.selectors, []);
    }
    const repoRoot = makeTempRepo();
    writeAzure(repoRoot, 'azure-pipelines.yml', { variables: { node: '$(node)' }, ...azureNode('$(node)') });
    const result = azureCli(repoRoot);
    assert.equal(result.status, 1);
    assert.match(result.report.problems[0].message, /cyclic/);
});

test('Azure template files enforce real containment while allowing contained directory links', () => {
    const repoRoot = makeTempRepo();
    const outside = makeTempRepo();
    writeAzure(outside, 'node.yml', azureNode('24'));
    writeAzure(repoRoot, 'inside/node.yml', azureNode('24'));
    for (const [name, target, succeeds] of [['external', outside, false], ['internal', path.join(repoRoot, 'inside'), true]]) {
        fs.symlinkSync(target, path.join(repoRoot, name), process.platform === 'win32' ? 'junction' : 'dir');
        writeAzure(repoRoot, 'azure-pipelines.yml', { steps: [{ template: `${name}/node.yml` }] });
        const result = azureCli(repoRoot);
        assert.equal(result.status, succeeds ? 0 : 1);
        if (succeeds) assert.deepEqual(result.report.problems, []);
        else assert(result.report.problems.some((item) => /escapes repository root/.test(item.message)));
    }
});

function azureMutant(repoRoot, replacements) {
    let source = fs.readFileSync(path.resolve(__dirname, '../../.github/scripts/check-toolchain-eol.js'), 'utf8');
    for (const [before, after, occurrences = 1] of replacements) {
        assert.equal(source.split(before).length - 1, occurrences, `Mutation anchor: ${before}`);
        source = source.split(before).join(after);
    }
    const file = path.join(repoRoot, 'mutant.js');
    fs.writeFileSync(file, source);
    return file;
}

test('Azure fixed CLI oracles detect traversal, binding and incomplete-result guard removal', () => {
    const fixtures = [
        ['traversal', [['return visit(selected, templateRole, argumentsMap, caller, budget, ancestry, discovery);', 'return new Map(caller.variables);']], '24', '18', 1, ['18']],
        ['binding', [['if (overrides !== null && !discovery)', 'if (false)']], '18', '24', 0, ['24']],
        ['problem', [['const addProblem = (source, error) => problems.push({ path: source, message: error.message });', 'const addProblem = () => {};']], '24', undefined, 1, []],
    ];
    for (const [name, replacements, defaultVersion, override, status, expected] of fixtures) {
        const repoRoot = makeTempRepo();
        writeAzure(repoRoot, 'azure-pipelines.yml', { extends: { template: name === 'problem' ? 'absent.yml' : 'ci/node.yml',
            parameters: { node: override } } });
        writeAzure(repoRoot, 'ci/node.yml', parameterizedAzure(defaultVersion));
        const fixed = (result) => {
            assert.equal(result.status, status, `${name} native failure oracle`);
            assert.deepEqual(result.report.selectors.map((item) => item.rawValue), expected, `${name} selected-value oracle`);
            if (name === 'problem') assert(result.report.problems.some((item) => /ENOENT/.test(item.message)));
            else assert.deepEqual(result.report.problems, []);
        };
        fixed(azureCli(repoRoot));
        assert.throws(() => fixed(azureCli(repoRoot, [], azureMutant(repoRoot, replacements))), assert.AssertionError);
    }
});

test('Azure independent containment oracle detects removal before reading an outside template', () => {
    const repoRoot = makeTempRepo();
    const outside = makeTempRepo();
    writeAzure(outside, 'node.yml', azureNode('24'));
    fs.symlinkSync(outside, path.join(repoRoot, 'linked'), process.platform === 'win32' ? 'junction' : 'dir');
    writeAzure(repoRoot, 'azure-pipelines.yml', { steps: [{ template: 'linked/node.yml' }] });
    const fixed = (result) => {
        assert.equal(result.status, 1);
        assert(result.report.problems.some((item) => /escapes repository root/.test(item.message)));
        assert.deepEqual(result.report.selectors, []);
    };
    fixed(azureCli(repoRoot));
    const mutant = azureMutant(repoRoot, [['assertPathWithinRepo(realRoot, real);', '// removed containment']]);
    assert.throws(() => fixed(azureCli(repoRoot, [], mutant)), assert.AssertionError);
});

for (const [kind, good, bad, diagnostic] of [
    ['files', 100, 101, /100-file/], ['depth', 100, 101, /100-level/],
    ['invocations', 4096, 4097, /4096-invocation/],
]) {
    test(`Azure ${kind} budget accepts its boundary and rejects overflow`, () => {
        for (const count of [good, bad]) {
            const repoRoot = makeTempRepo();
            if (kind === 'depth') {
                for (let index = 0; index < count; index += 1) {
                    writeAzure(repoRoot, index === 0 ? 'azure-pipelines.yml' : `ci/${index}.yml`, index + 1 === count ? { steps: [] }
                        : { extends: { template: index === 0 ? 'ci/1.yml' : `${index + 1}.yml` } });
                }
            } else {
                writeAzure(repoRoot, 'azure-pipelines.yml', { steps: Array.from({ length: count - 1 }, (_, index) => ({
                    template: `ci/${kind === 'files' ? index : 0}.yml`,
                })) });
                for (let index = 0; index < (kind === 'files' ? count - 1 : 1); index += 1) writeAzure(repoRoot, `ci/${index}.yml`, { steps: [] });
            }
            const result = azureCli(repoRoot);
            assert.equal(result.status, count === good ? 0 : 1);
            if (count === good) assert.deepEqual(result.report.problems, []);
            else assert(result.report.problems.some((item) => diagnostic.test(item.message)));
        }
    });
}

test('Azure byte and syntax-tree budgets have exact fixed boundaries', () => {
    for (const [bytes, succeeds] of [[2 * 1024 * 1024, true], [2 * 1024 * 1024 + 1, false]]) {
        const repoRoot = makeTempRepo();
        const prefix = 'steps: []\n#';
        writeFile(repoRoot, 'azure-pipelines.yml', prefix + 'x'.repeat(bytes - prefix.length));
        const result = azureCli(repoRoot);
        assert.equal(result.status, succeeds ? 0 : 1);
        if (succeeds) assert.deepEqual(result.report.problems, []);
        else assert(result.report.problems.some((item) => /2 MiB file/.test(item.message)));
    }
    for (const [count, succeeds] of [[199997, true], [199998, false]]) {
        const repoRoot = makeTempRepo();
        writeFile(repoRoot, 'azure-pipelines.yml', `steps: []\nunused: [${Array(count).fill('null').join(',')}]\n`);
        const result = azureCli(repoRoot);
        assert.equal(result.status, succeeds ? 0 : 1);
        if (succeeds) assert.deepEqual(result.report.problems, []);
        else assert(result.report.problems.some((item) => /200000-node/.test(item.message)));
    }
});

test('Azure cumulative bytes account for unique files and fail above twenty MiB', () => {
    for (const overflow of [false, true]) {
        const repoRoot = makeTempRepo();
        const files = Array.from({ length: 9 }, (_, index) => `ci/${index}.yml`);
        if (overflow) files.push('ci/overflow.yml');
        const entry = yaml.stringify({ steps: files.map((template) => ({ template })) }) + '#';
        writeFile(repoRoot, 'azure-pipelines.yml', entry + 'x'.repeat(2 * 1024 * 1024 - entry.length));
        for (const filename of files) {
            const prefix = 'steps: []\n#';
            writeFile(repoRoot, filename, filename.includes('overflow') ? 'steps: []\n' : prefix + 'x'.repeat(2 * 1024 * 1024 - prefix.length));
        }
        const result = azureCli(repoRoot);
        assert.equal(result.status, overflow ? 1 : 0);
        if (overflow) assert(result.report.problems.some((item) => /20 MiB cumulative/.test(item.message)));
        else assert.deepEqual(result.report.problems, []);
    }
});

test('Azure invocation-limit mutant fails a fixed overflow oracle', () => {
    const repoRoot = makeTempRepo();
    writeAzure(repoRoot, 'azure-pipelines.yml', { steps: Array.from({ length: 4096 }, () => ({ template: 'ci/empty.yml' })) });
    writeAzure(repoRoot, 'ci/empty.yml', { steps: [] });
    const fixed = (result) => {
        assert.equal(result.status, 1);
        assert(result.report.problems.some((item) => /4096-invocation/.test(item.message)));
    };
    fixed(azureCli(repoRoot));
    const mutant = azureMutant(repoRoot, [['budget.invocations > AZURE_INVENTORY_LIMITS.invocations', 'false']]);
    assert.throws(() => fixed(azureCli(repoRoot, [], mutant)), assert.AssertionError);
});

test('Azure inventory executes from an installed scratch tree without repository support files', () => {
    const repoRoot = makeTempRepo();
    const executable = path.join(repoRoot, 'check-toolchain-eol.js');
    fs.copyFileSync(path.resolve(__dirname, '../../.github/scripts/check-toolchain-eol.js'), executable);
    writeAzure(repoRoot, 'azure-pipelines.yml', { extends: { template: 'ci/build.yml' } });
    writeAzure(repoRoot, 'ci/build.yml', azureNode('24'));
    for (const absent of ['.template-sync', '.github/scripts/instruction_contract_core.py', '.github/instructions', 'pyproject.toml']) assert(!fs.existsSync(path.join(repoRoot, absent)));
    const result = azureCli(repoRoot, [], executable);
    assert.equal(result.status, 0);
    assert.deepEqual(result.report.problems, []);
    assert.deepEqual(result.report.selectors.map((item) => item.rawValue), ['24']);
});


test('Azure compile-time variables retain declaration scope across template calls', () => {
    for (const listed of [false, true]) {
        for (const [caller, callee] of [['18', '24'], ['24', '18']]) {
            const repoRoot = makeTempRepo();
            const declarations = { selected: '${{ parameters.node }}', alias: '${{ variables.selected }}',
                deferred: '$(runtime)', unrelated: '${{ unsupported.expression }}', literal: 'keep # meaningful text' };
            writeAzure(repoRoot, 'azure-pipelines.yml', {
                parameters: [{ name: 'node', type: 'string', default: caller }],
                variables: listed ? Object.entries(declarations).map(([name, value]) => ({ name, value })) : declarations,
                jobs: [{ job: 'first', variables: { runtime: '18' }, steps: [{ template: 'ci/build.yml' }] },
                    { job: 'second', variables: { runtime: '24' }, ...azureNode('$(deferred)') }],
            });
            writeAzure(repoRoot, 'ci/build.yml', {
                parameters: [{ name: 'node', type: 'string', default: callee }], ...azureNode('$(alias)'),
            });
            const fixed = (result) => {
                assert.equal(result.status, caller === '18' ? 1 : 0);
                assert.deepEqual(result.report.problems, []);
                assert.deepEqual(result.report.findings.map((item) => [item.path, item.rawValue, item.status]), [
                    ['ci/build.yml', caller, caller === '18' ? 'eol' : 'supported'],
                    ['azure-pipelines.yml', '24', 'supported'],
                ]);
            };
            fixed(azureCli(repoRoot));
            const mutant = azureMutant(repoRoot, [["if (!discovery && typeof value === 'string' &&", "if (false && typeof value === 'string' &&"]]);
            assert.throws(() => fixed(azureCli(repoRoot, [], mutant)), assert.AssertionError);
        }
    }
});

test('Azure unresolved compile-time declarations cannot capture callee defaults', () => {
    for (const selected of ['${{ parameters.absent }}', '${{ variables.absent }}', '${{ unsupported.expression }}']) {
        const repoRoot = makeTempRepo();
        writeAzure(repoRoot, 'azure-pipelines.yml', { variables: { selected }, steps: [{ template: 'ci/build.yml' }] });
        writeAzure(repoRoot, 'ci/build.yml', { parameters: [{ name: 'absent', default: '24' }],
            variables: { absent: '24' }, ...azureNode('$(selected)') });
        const result = azureCli(repoRoot);
        assert.equal(result.status, 1);
        assert(result.report.problems.length > 0);
        assert.deepEqual(result.report.selectors, []);
    }
});

function azureMatrixPipeline(matrix, variables = {}, steps = azureNode('$(node)').steps) {
    return { variables, jobs: [{ job: 'matrix', strategy: { matrix }, steps }] };
}

function assertAzureSelection(result, status, values, diagnostic) {
    assert.equal(result.status, status, 'Azure selector native status');
    assert.deepEqual(result.report.selectors.map((item) => item.rawValue), values,
        'Azure effective selector values');
    if (diagnostic) {
        assert(result.report.problems.some((item) => diagnostic.test(item.message)),
            'Azure selected-input diagnostic');
    } else {
        assert.deepEqual(result.report.problems, []);
    }
}

test('Azure matrix variants preserve missing and invalid selectors across all Node inputs', () => {
    for (const [task, input, extra] of [
        ['UseNode@1', 'version', {}],
        ['NodeTool@0', 'versionSpec', {}],
        ['NodeTool@0', 'versionFilePath', { versionSource: 'fromFile' }],
    ]) {
        for (const bad of [{}, { node: '' }, { node: '  ' }, { node: null },
            { node: ['24'] }, { node: { nested: '24' } }]) {
            const repoRoot = makeTempRepo();
            writeFile(repoRoot, '.nvmrc', '24\n');
            const good = input === 'versionFilePath' ? '.nvmrc' : '24';
            writeAzure(repoRoot, 'azure-pipelines.yml', azureMatrixPipeline(
                { good: { node: good }, bad }, {},
                [{ task, inputs: { ...extra, [input]: '$(node)' } }],
            ));
            assertAzureSelection(azureCli(repoRoot), 1, ['24'],
                /no checked-in value|nonblank checked-in scalar|matrix variable must be a checked-in scalar/);
        }
    }
});

test('Azure matrix selection preserves real fallback and discards fully shadowed values', () => {
    const cases = [
        [{ first: { node: '24' }, second: {} }, { node: '24' }, 0, ['24'], undefined],
        [{ first: { node: '24' }, second: { node: '' } }, { node: '24' }, 1, ['24'], /nonblank/],
        [{ first: { node: '24' }, second: { node: null } }, { node: '24' }, 1, ['24'], /no checked-in value/],
        [{ first: { node: '24' }, second: { node: '24' } }, { node: '18' }, 0, ['24'], undefined],
        [{ first: { node: '24' }, second: {} }, { node: '18' }, 1, ['24', '18'], undefined],
    ];
    for (const [matrix, variables, status, expected, diagnostic] of cases) {
        const repoRoot = makeTempRepo();
        writeAzure(repoRoot, 'azure-pipelines.yml', azureMatrixPipeline(matrix, variables));
        assertAzureSelection(azureCli(repoRoot), status, expected, diagnostic);
    }
});

test('Azure matrix resolution retains supported compile-time variable syntax', () => {
    const repoRoot = makeTempRepo();
    writeAzure(repoRoot, 'azure-pipelines.yml', azureMatrixPipeline({
        good: { node: '24' }, bad: {},
    }, {}, azureNode('${{ variables.node }}').steps));
    assertAzureSelection(azureCli(repoRoot), 1, ['24'], /no checked-in value/);
});

test('Azure matrix aliases stay correlated to their own variants and included templates', () => {
    for (const viaTemplate of [false, true]) {
        const repoRoot = makeTempRepo();
        writeAzure(repoRoot, 'ci/node.yml', azureNode('$(selected)'));
        const steps = viaTemplate ? [{ template: 'ci/node.yml' }] : azureNode('$(selected)').steps;
        writeAzure(repoRoot, 'azure-pipelines.yml', azureMatrixPipeline({
            first: { selected: '$(node)', node: '24' },
            second: { selected: '$(node)' },
        }, {}, steps));
        assertAzureSelection(azureCli(repoRoot), 1, ['24'], /no checked-in value/);
        writeAzure(repoRoot, 'azure-pipelines.yml', azureMatrixPipeline({
            first: { selected: '$(node)', node: '24' },
            second: { selected: '24' },
        }, {}, steps));
        assertAzureSelection(azureCli(repoRoot), 0, ['24']);
    }
});

test('Azure selected matrix shapes fail while literal inputs ignore unrelated matrices', () => {
    for (const matrix of ['$[ dependencies.generator.outputs.matrix ]', [], null, {}]) {
        const repoRoot = makeTempRepo();
        writeAzure(repoRoot, 'azure-pipelines.yml', azureMatrixPipeline(matrix, { node: '24' }));
        assertAzureSelection(azureCli(repoRoot), 1, [], /matrix must be a nonempty checked-in mapping/);
        writeAzure(repoRoot, 'azure-pipelines.yml', azureMatrixPipeline(matrix, {}, azureNode('24').steps));
        assertAzureSelection(azureCli(repoRoot), 0, ['24']);
    }
    for (const bad of [null, '24', ['24']]) {
        const repoRoot = makeTempRepo();
        writeAzure(repoRoot, 'azure-pipelines.yml',
            azureMatrixPipeline({ good: { node: '24' }, bad }));
        assertAzureSelection(azureCli(repoRoot), 1, ['24'], /matrix leg must be a checked-in variable mapping/);
    }
});

test('Azure parameter alternatives retain explicit blank defaults and values', () => {
    for (const parameter of [
        { name: 'node', default: '24', values: ['24', ''] },
        { name: 'node', default: '', values: ['24'] },
        { name: 'node', default: '24', values: ['24', '  '] },
        { name: 'node', default: null, values: ['24'] },
        { name: 'node', default: '24', values: ['24', null] },
        { name: 'node', default: '24', values: ['24', { nested: '24' }] },
        { name: 'node', default: '24', values: ['24', ['24']] },
        { name: 'node', default: ['24'], values: ['24'] },
    ]) {
        const repoRoot = makeTempRepo();
        const pipeline = parameterizedAzure('24');
        pipeline.parameters = [parameter];
        writeAzure(repoRoot, 'azure-pipelines.yml', pipeline);
        assertAzureSelection(azureCli(repoRoot), 1, ['24'], /nonblank checked-in scalar|must resolve to a checked-in scalar/);
    }
    const repoRoot = makeTempRepo();
    const pipeline = parameterizedAzure('24');
    pipeline.parameters = [{ name: 'node', values: ['24'] }];
    writeAzure(repoRoot, 'azure-pipelines.yml', pipeline);
    assertAzureSelection(azureCli(repoRoot), 0, ['24']);
});

test('Azure selected-input empty checks preserve unrelated arguments and actual template bindings', () => {
    const repoRoot = makeTempRepo();
    writeAzure(repoRoot, 'azure-pipelines.yml', { extends: { template: 'ci/node.yml',
        parameters: { node: '24', unrelated: '' } } });
    const template = parameterizedAzure('18');
    template.parameters[0].values = ['18', ''];
    template.parameters.push({ name: 'unrelated', type: 'string' });
    writeAzure(repoRoot, 'ci/node.yml', template);
    assertAzureSelection(azureCli(repoRoot), 0, ['24']);
    writeAzure(repoRoot, 'azure-pipelines.yml', { extends: { template: 'ci/node.yml',
        parameters: { node: '', unrelated: '' } } });
    assertAzureSelection(azureCli(repoRoot), 1, [], /nonblank checked-in scalar/);
});

test('Azure matrix map copying consumes the existing node work budget', () => {
    for (const count of [400, 500]) {
        const repoRoot = makeTempRepo();
        const variables = Object.fromEntries(Array.from({ length: 400 }, (_, i) => ['unused' + i, 'x']));
        const matrix = Object.fromEntries(Array.from({ length: count }, (_, i) => ['leg' + i, { node: '24' }]));
        writeAzure(repoRoot, 'azure-pipelines.yml', azureMatrixPipeline(matrix, variables));
        const fixed = (result) => assertAzureSelection(result, count === 400 ? 0 : 1, count === 400 ? ['24'] : [],
            count === 400 ? undefined : /200000-node work limit/);
        fixed(azureCli(repoRoot));
        if (count === 500) {
            const mutant = azureMutant(repoRoot, [[
                'context.inventoryBudget.nodes += context.variables.size + Object.keys(entry).length;',
                '// mutation: map copying is no longer charged',
            ]]);
            assert.throws(() => fixed(azureCli(repoRoot, [], mutant)), assert.AssertionError);
        }
    }
});

test('Azure fixed selector oracles detect isolated matrix and parameter guard mutations', () => {
    const cases = [
        ['matrix', [
            ["if (context.strategy && Object.hasOwn(context.strategy, 'matrix'))", 'if (false)'],
        ], azureMatrixPipeline({ good: { node: '24' }, bad: { node: '' } }, { node: '24' }),
        1, ['24'], /nonblank/],
        ['missing', [
            ['if (values.length === 0) return fail', 'if (values.length === 0) return []; // removed diagnostic\n    // return fail'],
        ], azureMatrixPipeline({ good: { node: '24' }, bad: {} }),
        1, ['24'], /no checked-in value/],
        ['override', [
            ['variables: mergeMaps(context.variables, new Map(Object.entries(entry))),', 'variables: context.variables,'],
        ], azureMatrixPipeline({ good: { node: '24' } }, { node: '18' }),
        0, ['24'], undefined],
        ['blank', [
            ["if (typeof resolved.rawValue === 'string' && !resolved.rawValue.trim())", 'if (false)'],
        ], { ...parameterizedAzure('24'), parameters: [{ name: 'node', default: '24', values: ['24', ''] }] },
        1, ['24'], /nonblank/],
        ['shape', [
            ["return fail('Azure selected variable matrix must be a nonempty checked-in mapping.');", 'return [];'],
        ], azureMatrixPipeline('$[ dependencies.generator.outputs.matrix ]', { node: '24' }),
        1, [], /matrix must be a nonempty checked-in mapping/],
        ['leg', [
            ["fail('Azure selected matrix leg must be a checked-in variable mapping.');", '// mutation: incomplete leg ignored'],
        ], azureMatrixPipeline({ good: { node: '24' }, bad: null }),
        1, ['24'], /matrix leg must be a checked-in variable mapping/],
        ['scalar', [
            ["return fail('Azure selected matrix variable must be a checked-in scalar.');", '// mutation: structural variable allowed'],
        ], azureMatrixPipeline({ good: { node: '24' }, bad: { node: ['24'] } }),
        1, ['24'], /matrix variable must be a checked-in scalar/],
        ['alternatives', [
            ['if (values.length === 0) return fail', 'values = uniqueValues(values);\n    if (values.length === 0) return fail'],
        ], { ...parameterizedAzure('24'), parameters: [{ name: 'node', default: '24', values: ['24', ''] }] },
        1, ['24'], /nonblank/],
        ['parameter-scalar', [
            ['if (parameterMatch && Array.isArray(rawValue))', 'if (false)'],
        ], { ...parameterizedAzure('24'), parameters: [{ name: 'node', default: '24', values: ['24', ['24']] }] },
        1, ['24'], /parameter selector alternative must resolve to a checked-in scalar/],
    ];
    for (const [name, replacements, pipeline, status, expected, diagnostic] of cases) {
        const repoRoot = makeTempRepo();
        writeAzure(repoRoot, 'azure-pipelines.yml', pipeline);
        const fixed = (result) => assertAzureSelection(result, status, expected, diagnostic);
        fixed(azureCli(repoRoot));
        assert.throws(() => fixed(azureCli(repoRoot, [], azureMutant(repoRoot, replacements))),
            assert.AssertionError, name);
    }
});
