import { config as n8nConfig } from '@n8n/node-cli/eslint';

/** @type {import('eslint').Linter.Config[]} */
export default [
	...n8nConfig,
	{
		ignores: [
			'dist/**',
			'node_modules/**',
			'**/*.test.ts',
			'*.js',
			'*.cjs',
			'*.mjs',
			'coverage/**',
			'.coverage/**',
		],
	},
	{
		files: ['package.json'],
		rules: {
			// This node is derived from n8n's Sustainable Use Licensed source, so the
			// package cannot claim MIT. See LICENSE.md.
			'n8n-nodes-base/community-package-json-license-not-default': 'off',
		},
	},
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
		files: ['nodes/*/descriptions/**/*.ts'],
		rules: {
			// These are declarative parameter definitions, not logic: they are long
			// because the Graph surface is large, and splitting them by line count
			// would fragment a single node's parameters across arbitrary files.
			'max-lines': 'off',
			// The postReceive handlers issue follow-up Graph writes per item. Those must
			// stay sequential: Graph throttles aggressively, and the create handlers roll
			// back the just-created object before moving on to the next item.
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
