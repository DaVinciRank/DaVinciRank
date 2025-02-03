import babelPresetEnv from "@babel/preset-env";
import babelPresetTypescript from "@babel/preset-typescript";
import babelPluginTransformRuntime from "@babel/plugin-transform-runtime";

export default {
  presets: [
    [
      babelPresetEnv,
      {
        targets: {
          node: "current",
        },
      },
    ],
    [babelPresetTypescript],
  ],
  plugins: [
    [babelPluginTransformRuntime],
  ],
};