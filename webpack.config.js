const devCerts = require("office-addin-dev-certs");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const httpsOptions = dev ? await devCerts.getHttpsServerOptions() : {};

  return {
    mode: dev ? "development" : "production",

    entry: {
        taskpane: "./src/taskpane/taskpane.ts",
    },

    output: {
      clean: true,
      filename: "[name].js",
    },

    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
    },

    module: {
      rules: [
        {
            test: /\.ts$/,
            exclude: /node_modules/,
            use: "ts-loader",
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
      ],
    },

    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
        }),

      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets",
            to: "assets",
            noErrorOnMissing: true,
          },
          {
            from: "docs/support.html",
            to: "support.html",
            noErrorOnMissing: true,
          },
        ],
      }),
    ],

    devServer: {
      port: 3000,
      server: {
        type: "https",
        options: httpsOptions,
      },
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
    },
  };
};