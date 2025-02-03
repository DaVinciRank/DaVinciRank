import glob from "glob";
import typescript from "rollup-plugin-typescript2";
import cleanup from "rollup-plugin-cleanup";
import license from "rollup-plugin-license";

import packageJson from "./package.json" with { type: "json" };

export default {
  input: glob.sync("src/**/*.ts"),
  // input: "src/main.ts",
  output: {
    dir: "dist",
    format: "esm",
  },
  plugins: [
    cleanup({ comments: "none", extensions: [".ts"] }),
    license({
      banner: [
        `Name: ${packageJson.name}`,
        `Version: ${packageJson.version}`,
        `Description: ${packageJson.description}`,
        `Author: ${packageJson.author}`,
        `@see ${packageJson.homepage}`,
      ].join("\n"),
    }),
    typescript(),
  ],
  context: "this",
};
