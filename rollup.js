const commonjs = require("rollup-plugin-commonjs");
const typescript = require("rollup-plugin-typescript2");
const nodeResolve = require("rollup-plugin-node-resolve");
const inject = require("@rollup/plugin-inject");
const rollup = require("rollup");
const path = require("path");

async function build() {
    const bundle = await rollup.rollup({
        input: "src/index.ts",
        external: ["fs", "path", "os", "crypto", "stream", "util"],
        plugins: [
            commonjs(),
            nodeResolve({ include: "node_modules/**" }),
            typescript({
                tsconfigOverride: {
                    compilerOptions: { module: "es2015" },
                },
            }),
            inject({}),
        ],
    });
    await bundle.write({
        file: "dist/index.js",
        format: "cjs",
        globals: {
            fs: "fs",
            path: "path",
            os: "os",
            crypto: "crypto",
            stream: "stream",
            util: "util",
        },
    });
}
build();
