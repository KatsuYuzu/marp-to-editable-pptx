const path = require('node:path')
const esbuild = require('esbuild')
const { EsbuildPlugin } = require('esbuild-loader')
const { NormalModuleReplacementPlugin, IgnorePlugin } = require('webpack')

module.exports = (env) => ({
  mode: env?.production ? 'production' : 'none',
  target: 'node',
  entry: './src/extension.ts',
  output: {
    filename: 'extension.js',
    path: path.resolve(__dirname, 'dist'),
    clean: !!env?.production,
    library: { type: 'commonjs' },
  },
  resolve: {
    extensions: ['.ts', '.js'],
  },
  module: {
    rules: [
      {
        test: /\.ts$/,
        loader: 'esbuild-loader',
        options: {
          implementation: esbuild,
          target: 'es2021',
        },
      },
    ],
  },
  externals: {
    vscode: 'commonjs vscode',
  },
  optimization: {
    minimizer: [
      new EsbuildPlugin({
        target: 'es2021',
        format: 'cjs',
        keepNames: true,
      }),
    ],
  },
  plugins: [
    // Remove 'node:' prefix so webpack can resolve built-in modules
    new NormalModuleReplacementPlugin(/^node:/, (resource) => {
      resource.request = resource.request.replace(/^node:/, '')
    }),
    // 'emitter' is a missing transitive dep of serve-index (used only in
    // marp-cli's server/preview mode). We never use that code path, so it
    // is safe to ignore at bundle time.
    new IgnorePlugin({ resourceRegExp: /^emitter$/ }),
  ],
  performance: { hints: false },
  devtool: env?.production ? false : 'nosources-source-map',
})
