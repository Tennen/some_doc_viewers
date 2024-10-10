import { defineConfig } from "tsup";
import { lessLoader } from 'esbuild-plugin-less';

export default defineConfig({
  entry: ["src/index.ts"],
  clean: true,
  format: ["cjs", "esm"],
  esbuildPlugins: [lessLoader()],
  esbuildOptions: (options) => {
    options.ignoreAnnotations = true
  },
  loader: {
    '.less': 'css',
  },
});
