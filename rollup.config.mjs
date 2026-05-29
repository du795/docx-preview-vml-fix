import typescript from '@rollup/plugin-typescript';
import terser from '@rollup/plugin-terser';

const output = { // 通用输出配置
	banner: `/*
 * @license
 * docx-preview <https://github.com/VolodymyrBaydalka/docxjs>
 * Released under Apache License 2.0  <https://github.com/VolodymyrBaydalka/docxjs/blob/master/LICENSE>
 * Copyright Volodymyr Baydalka
 */`,
	sourcemap: true,
}

const umdOutput = { // UMD 构建产物配置
	...output,
	name: "docx",
	file: 'dist/docx-preview.js',
	format: 'umd',
	globals: {
		jszip: 'JSZip'
	},
};

const esOutput = { // ESM 构建产物配置
	...output,
	file: 'dist/docx-preview.mjs',
	format: 'es',
};

export default args => {
	const isProductionBuild = args.environment == 'BUILD:production'; // 是否生产构建
	const config = { // Rollup 构建配置
		input: 'src/docx-preview.ts',
		output: isProductionBuild ? [umdOutput,
			{
				...umdOutput,
				file: 'dist/docx-preview.min.js',
				plugins: [terser()]
			},
			esOutput,
			{
				...esOutput,
				file: 'dist/docx-preview.min.mjs',
				plugins: [terser()]
			}] : [umdOutput, esOutput],
		plugins: [typescript()]
	}

	return config
};
