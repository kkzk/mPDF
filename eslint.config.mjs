import tseslint from 'typescript-eslint';

export default tseslint.config(
	{ ignores: ['out', 'dist', '**/*.d.ts'] },
	...tseslint.configs.recommended,
	{
		files: ['src/**/*.ts'],
		rules: {
			'@typescript-eslint/naming-convention': 'warn',
			curly: 'warn',
			eqeqeq: 'warn',
			'no-throw-literal': 'warn',
			semi: 'warn',
		},
	},
);
