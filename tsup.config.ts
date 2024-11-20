import { defineConfig } from "tsup";
import { lessLoader } from 'esbuild-plugin-less';

export default defineConfig({
  entry: ["src/index.ts"],
  clean: true,
  splitting: true,
  format: ["esm", "cjs"],
  esbuildPlugins: [lessLoader()],
  esbuildOptions: (options) => {
    options.ignoreAnnotations = true
  },
  loader: {
    '.less': 'css',
  },
  minify: true
});
