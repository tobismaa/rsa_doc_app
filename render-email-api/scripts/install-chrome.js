const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const rootDir = path.resolve(__dirname, '..');
const cacheDir = process.env.PUPPETEER_CACHE_DIR || path.join(rootDir, '.cache', 'puppeteer');
process.env.PUPPETEER_CACHE_DIR = cacheDir;
fs.mkdirSync(cacheDir, { recursive: true });

const binName = process.platform === 'win32' ? 'puppeteer.cmd' : 'puppeteer';
const localBin = path.join(rootDir, 'node_modules', '.bin', binName);
const command = fs.existsSync(localBin)
    ? localBin
    : (process.platform === 'win32' ? 'npx.cmd' : 'npx');
const args = fs.existsSync(localBin)
    ? ['browsers', 'install', 'chrome']
    : ['puppeteer', 'browsers', 'install', 'chrome'];

const result = spawnSync(command, args, {
    cwd: rootDir,
    stdio: 'inherit',
    env: process.env
});

if (result.status !== 0) {
    process.exit(result.status || 1);
}
