import { spawn } from "node:child_process";
import process from "node:process";

const buildCommand = process.platform === "win32" ? "npm.cmd" : "npm";
const buildArgs = ["run", "build", "--", "--watch", "--preserveWatchOutput"];
const dockerCommand = process.platform === "win32" ? "docker.exe" : "docker";
const dockerArgs = [
  "compose",
  "-f",
  "../n8n/docker-compose.wsl.yml",
  "restart",
  "n8n",
];

let restartInFlight = false;
let restartQueued = false;
let shuttingDown = false;

const buildProcess = spawn(buildCommand, buildArgs, {
  cwd: process.cwd(),
  stdio: ["inherit", "pipe", "pipe"],
});

const forwardOutput = (stream, target) => {
  stream.setEncoding("utf8");
  stream.on("data", (chunk) => {
    target.write(chunk);
    if (chunk.includes("Found 0 errors. Watching for file changes.")) {
      scheduleRestart();
    }
  });
};

const scheduleRestart = () => {
  if (shuttingDown) {
    return;
  }

  if (restartInFlight) {
    restartQueued = true;
    return;
  }

  restartInFlight = true;
  const restartProcess = spawn(dockerCommand, dockerArgs, {
    cwd: process.cwd(),
    stdio: "inherit",
  });

  restartProcess.on("exit", (code) => {
    restartInFlight = false;
    if (code !== 0) {
      console.error(
        `Docker restart failed with exit code ${code ?? "unknown"}.`,
      );
    }

    if (restartQueued) {
      restartQueued = false;
      scheduleRestart();
    }
  });

  restartProcess.on("error", (error) => {
    restartInFlight = false;
    console.error("Failed to start Docker restart command.", error);
    if (restartQueued) {
      restartQueued = false;
      scheduleRestart();
    }
  });
};

const shutdown = (signal) => {
  if (shuttingDown) {
    return;
  }

  shuttingDown = true;
  buildProcess.kill(signal);
};

forwardOutput(buildProcess.stdout, process.stdout);
forwardOutput(buildProcess.stderr, process.stderr);

buildProcess.on("exit", (code, signal) => {
  if (signal) {
    process.exit(0);
  }

  process.exit(code ?? 0);
});

buildProcess.on("error", (error) => {
  console.error("Failed to start TypeScript watch build.", error);
  process.exit(1);
});

process.on("SIGINT", () => shutdown("SIGINT"));
process.on("SIGTERM", () => shutdown("SIGTERM"));
