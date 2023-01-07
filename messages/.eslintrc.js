require("@rushstack/eslint-config/patch/modern-module-resolution");
module.exports = {
  extends: ["@microsoft/eslint-config-spfx/lib/profiles/react"],
  parserOptions: { tsconfigRootDir: __dirname },
  rules: {
    "@typescript-eslint/no-empty-interface": "off",
    "@typescript-eslint/no-inferrable-types": "off",
  },
};
