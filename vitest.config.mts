import { defineConfig } from 'vitest/config';

export default defineConfig({
	test: {
		globals: true,
		environment: 'node',
		include: ['**/*.test.ts'],
		exclude: ['node_modules/**', 'dist/**'],
		coverage: {
			provider: 'v8',
			// Dot-prefixed: `n8n-node build` copies `**/*.{png,svg}` from the whole project
			// into dist, and the HTML report ships PNGs. fast-glob skips dot-directories,
			// so this keeps the report out of the published artifact.
			reportsDirectory: '.coverage',
			reporter: ['text', 'html', 'lcov'],
			include: ['credentials/**/*.ts', 'nodes/**/*.ts'],
			exclude: ['**/*.test.ts', '**/__schema__/**'],
			thresholds: {
				lines: 80,
				functions: 80,
				branches: 80,
				statements: 80,
			},
		},
	},
});
