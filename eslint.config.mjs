import { config as n8nConfig } from '@n8n/node-cli/eslint';

/** @type {import('eslint').Linter.Config[]} */
export default [
	...n8nConfig,
	{
		// Test files are deliberately NOT ignored. n8n's verification scanner lints
		// everything under `nodes/` and `credentials/` — tests included — so excluding
		// them here would hide gate failures until submission.
		ignores: [
			'dist/**',
			'node_modules/**',
			'*.js',
			'*.cjs',
			'*.mjs',
			'coverage/**',
			'.coverage/**',
		],
	},
	// NOTE: the MIT-license rule is deliberately left at its default. It used to be
	// relaxed here because the package was on n8n's Sustainable Use License and could not
	// satisfy it; the package is MIT from 1.0.1, so the rule passes on its own and nothing
	// needs suppressing. n8n's scanner builds its own config and never reads this file, so
	// a local override would not have affected the verification gate anyway.
	{
		files: ['**/*.ts'],
		rules: {
			'@typescript-eslint/no-explicit-any': 'warn',
			'@typescript-eslint/no-unused-vars': [
				'error',
				{
					argsIgnorePattern: '^_',
					varsIgnorePattern: '^_',
					caughtErrorsIgnorePattern: '^_',
				},
			],
			complexity: ['warn', { max: 15 }],
			'max-lines-per-function': [
				'warn',
				{
					max: 100,
					skipBlankLines: true,
					skipComments: true,
				},
			],
			'max-lines': [
				'warn',
				{
					max: 400,
					skipBlankLines: true,
					skipComments: true,
				},
			],
			'max-params': ['warn', { max: 5 }],
			'max-depth': ['warn', { max: 4 }],
			'max-nested-callbacks': ['warn', { max: 3 }],
			'prefer-const': 'error',
			'no-var': 'error',
			eqeqeq: ['error', 'always'],
			'no-console': [
				'warn',
				{
					allow: ['warn', 'error', 'info'],
				},
			],
			'no-debugger': 'error',
			'no-duplicate-imports': 'error',
			'no-eval': 'error',
			'no-implied-eval': 'error',
			'no-new-func': 'error',
			curly: ['error', 'all'],
			'default-case': 'warn',
			'no-else-return': ['warn', { allowElseIf: false }],
			'prefer-template': 'warn',
			'object-shorthand': ['warn', 'always'],
			'arrow-body-style': ['warn', 'as-needed'],
			'prefer-arrow-callback': 'warn',
			'no-useless-return': 'warn',
			'no-useless-concat': 'warn',
			'prefer-destructuring': [
				'warn',
				{
					array: false,
					object: true,
				},
			],
			'no-unreachable': 'error',
			'no-constant-condition': 'error',
			'no-duplicate-case': 'error',
			'no-empty': ['error', { allowEmptyCatch: true }],
			'no-extra-semi': 'error',
			'no-irregular-whitespace': 'error',
			'valid-typeof': 'error',
			'no-cond-assign': ['error', 'always'],
			'no-await-in-loop': 'warn',
			'no-return-await': 'warn',
			'prefer-promise-reject-errors': 'error',
		},
	},
	// Placed after the general `**/*.ts` block so these stay switched off — in flat
	// config the later matching entry wins.
	{
		files: ['nodes/*/actions/**/*.operation.ts'],
		rules: {
			// An operation file is mostly its parameter definitions, which are long because
			// the Graph surface is large. Splitting one by line count would separate an
			// operation's parameters from the `execute` that reads them.
			'max-lines': 'off',
		},
	},
	{
		files: ['nodes/*/actions/router.ts'],
		rules: {
			// Awaiting inside the loop is the point: operations run one item at a time, so
			// that a directory write finishes before the next one starts.
			'no-await-in-loop': 'off',
		},
	},
	{
		files: ['**/*.test.ts', '**/*.spec.ts', 'tests/**/*.ts'],
		rules: {
			'@typescript-eslint/no-explicit-any': 'error',
			'max-lines-per-function': 'off',
			'max-lines': 'off',
			'@n8n/community-nodes/no-restricted-imports': 'off',
			'@n8n/community-nodes/no-restricted-globals': 'off',
		},
	},
	{
		files: ['*.config.ts', '*.config.mjs', '*.config.js'],
		rules: {
			'@n8n/community-nodes/no-restricted-imports': 'off',
		},
	},
];
