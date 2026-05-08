import { spawn } from "node:child_process";
import { copyFile, mkdir, readdir, rm } from "node:fs/promises";
import path from "node:path";
import process from "node:process";

const rootDir = process.cwd();
const distDir = path.join(rootDir, "dist");
const nodesDir = path.join(rootDir, "nodes");
const tscCommand = process.platform === "win32" ? "tsc.cmd" : "tsc";
const tscArgs = ["-p", "tsconfig.json", ...process.argv.slice(2)];
const watchMode = process.argv.includes("--watch");

async function cleanDist() {
  await rm(distDir, { recursive: true, force: true });
}

async function copySvgFiles(sourceDir) {
  const entries = await readdir(sourceDir, { withFileTypes: true });

  for (const entry of entries) {
    const sourcePath = path.join(sourceDir, entry.name);

    if (entry.isDirectory()) {
      await copySvgFiles(sourcePath);
      continue;
    }

    if (!entry.isFile() || !entry.name.endsWith(".svg")) {
      continue;
    }

    const relativePath = path.relative(rootDir, sourcePath);
    const destinationPath = path.join(distDir, relativePath);

    await mkdir(path.dirname(destinationPath), { recursive: true });
    await copyFile(sourcePath, destinationPath);
  }
}

async function runTsc() {
  await new Promise((resolve, reject) => {
    const child = spawn(tscCommand, tscArgs, {
      cwd: rootDir,
      stdio: "inherit",
    });

    child.on("error", reject);
    child.on("exit", (code) => resolve(code ?? 0));
  }).then(async (code) => {
    if (!watchMode && code === 0) {
      await copySvgFiles(nodesDir);
    }

    process.exit(code);
  });
}

try {
  await cleanDist();

  if (watchMode) {
    await copySvgFiles(nodesDir);
  }

  await runTsc();
} catch (error) {
  console.error("Build failed.", error);
  process.exit(1);
}
