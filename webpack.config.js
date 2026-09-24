/* eslint-disable no-undef */
const path = require("path");
const devCerts = require("office-addin-dev-certs");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

const DEV_URL = "https://localhost:3000/";
// Where the built add-in is hosted in production, e.g. ADDIN_URL=https://bangla-addin.vercel.app/
const PROD_URL = process.env.ADDIN_URL || DEV_URL;

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const baseUrl = PROD_URL.endsWith("/") ? PROD_URL : PROD_URL + "/";

  return {
    devtool: dev ? "source-map" : false,
    entry: {
      taskpane: "./src/taskpane/taskpane.ts",
      functions: "./src/functions/functions.ts",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      clean: true,
    },
    resolve: { extensions: [".ts", ".js"] },
    module: {
      rules: [{ test: /\.ts$/, use: "ts-loader", exclude: /node_modules/ }],
    },
    plugins: [
      // Shared runtime: the task pane page also hosts the custom functions.
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["functions", "taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets", to: "assets" },
          { from: "src/taskpane/taskpane.css", to: "taskpane.css" },
          { from: "src/functions/functions.json", to: "functions.json" },
          {
            from: "manifest.xml",
            to: "manifest.xml",
            transform(content) {
              return dev ? content : content.toString().split(DEV_URL).join(baseUrl);
            },
          },
        ],
      }),
    ],
    devServer: {
      headers: { "Access-Control-Allow-Origin": "*" },
      server: { type: "https", options: dev ? await devCerts.getHttpsServerOptions() : {} },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };
};
