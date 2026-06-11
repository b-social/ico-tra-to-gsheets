export default [
  {
    files: ["**/*.gs", "**/*.js"],
    languageOptions: {
      ecmaVersion: 2017,   // Apps Script is ES5/ES2017 compat
      sourceType: "script",
      globals: {
        // Google Apps Script globals
        SpreadsheetApp: "readonly",
        Logger:         "readonly",
        Session:        "readonly",
        Browser:        "readonly",
      },
    },
    rules: {
      "no-unused-vars":        ["warn", { vars: "all", args: "after-used" }],
      "no-undef":              "error",
      "no-unreachable":        "error",
      "no-constant-condition": "error",
      "eqeqeq":                ["error", "always"],
      "no-implicit-globals":   "off",   // GAS uses global functions
      "no-redeclare":          "error",
    },
  },
];
