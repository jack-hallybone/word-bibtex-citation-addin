import js from "@eslint/js";
import globals from "globals";
import prettier from "eslint-config-prettier";
import tseslint from "typescript-eslint";

export default tseslint.config(
  {
    // public/ is upstream bytes, redistributed unmodified: formatting it is a licence question, not a style one.
    ignores: ["dist/**", "public/**", "node_modules/**"],
  },

  js.configs.recommended,

  {
    files: ["src/**/*.ts", "tests/**/*.ts", "vite.config.ts"],
    extends: [tseslint.configs.recommendedTypeChecked],
    languageOptions: {
      parserOptions: {
        project: ["./tsconfig.json", "./tsconfig.tests.json"],
        tsconfigRootDir: import.meta.dirname,
      },
    },
    rules: {
      // Load-bearing: console.log would put bibliography/document content into a host log the add-in's CSP cannot reach.
      "no-console": ["error", { allow: ["warn", "error"] }],
      // Load-bearing: Error's cause option is ES2022, but src/ is held to the ES2019 webview floor.
      "preserve-caught-error": "off",
      // Load-bearing: the leading `_` marks a parameter kept only to match a host signature.
      "@typescript-eslint/no-unused-vars": [
        "error",
        {
          argsIgnorePattern: "^_",
          varsIgnorePattern: "^_",
          caughtErrorsIgnorePattern: "^_",
        },
      ],
    },
  },

  {
    // The suite deliberately builds malformed values and partial host objects, so the unsafe-* rules would just fight the mocks.
    files: ["tests/**/*.ts"],
    rules: {
      // node:test's test() owns its own promise; exempted only there — a real floating promise elsewhere is how a test assertion silently never runs.
      "@typescript-eslint/no-floating-promises": [
        "error",
        {
          allowForKnownSafeCalls: [
            {
              from: "package",
              package: "node:test",
              name: ["test", "it", "describe"],
            },
          ],
        },
      ],
      "@typescript-eslint/no-explicit-any": "off",
      "@typescript-eslint/no-unsafe-assignment": "off",
      "@typescript-eslint/no-unsafe-member-access": "off",
      "@typescript-eslint/no-unsafe-argument": "off",
      "@typescript-eslint/no-unsafe-call": "off",
      "@typescript-eslint/no-unsafe-return": "off",
    },
  },

  {
    // Plain Node ESM, outside any tsconfig, so type-aware rules read nothing.
    files: ["**/*.mjs"],
    extends: [tseslint.configs.disableTypeChecked],
    languageOptions: {
      globals: globals.node,
      sourceType: "module",
    },
    rules: {
      "no-console": "off",
    },
  },

  {
    // src/safe-json.js is plain JS, so it's covered here rather than by a config of its own.
    files: ["src/**/*.ts", "src/**/*.js"],
    languageOptions: { globals: globals.browser },
  },

  // Last, so no rule here can disagree with the formatter.
  prettier,
);
