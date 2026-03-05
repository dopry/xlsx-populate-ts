import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    globals: true,
    projects: [
      {
        test: {
          name: "unit",
          globals: true,
          include: ["test/unit/**/*.spec.ts"],
          setupFiles: ["test/helpers/matchers.ts"],
        },
      },
      {
        test: {
          name: "e2e-parse",
          globals: true,
          include: ["test/e2e-parse/**/*.spec.ts"],
          setupFiles: ["test/helpers/matchers.ts"],
          testTimeout: 30000,
        },
      },
      {
        test: {
          name: "e2e-generate",
          globals: true,
          include: ["test/e2e-generate/**/*.spec.ts"],
          setupFiles: ["test/helpers/matchers.ts"],
          testTimeout: 60000,
        },
      },
    ],
  },
});
