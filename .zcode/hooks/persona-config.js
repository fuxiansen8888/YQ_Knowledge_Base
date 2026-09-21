#!/usr/bin/env node
// persona-config — caveman/ponytail 共享的模式解析与读写。
//
// 解析顺序（每个名字独立）：
//   1. <NAME>_DEFAULT_MODE 环境变量
//   2. 仓库本地配置：从 cwd 向上找 .<name>/config.json 或 .<name>.json
//   3. 用户配置：~/.config/<name>/config.json
//   4. 默认值 'full'
//
// CLI:
//   node persona-config.js get <name>                 # 打印当前解析到的模式
//   node persona-config.js set <name> <mode> [dir]    # 写仓库本地配置（默认 cwd）

const fs = require('fs');
const path = require('path');
const os = require('os');

const MODES = {
  caveman: ['off', 'lite', 'full', 'ultra'],
  ponytail: ['off', 'lite', 'full', 'ultra'],
};
const DEFAULT_MODE = 'full';

function readJson(p) {
  try {
    return JSON.parse(fs.readFileSync(p, 'utf8'));
  } catch (e) {
    return null;
  }
}

// 从 startDir 向上找仓库本地配置文件，拒绝符号链接（与 caveman-config.js 同策略）。
function findRepoConfig(startDir, name) {
  let dir = startDir;
  for (let i = 0; i < 64; i++) {
    for (const rel of [`.${name}/config.json`, `.${name}.json`]) {
      const p = path.join(dir, rel);
      try {
        const st = fs.lstatSync(p);
        if (st.isFile() && !st.isSymbolicLink()) return p;
      } catch (e) { /* 不存在，试下一个 */ }
    }
    const parent = path.dirname(dir);
    if (parent === dir) return null;
    dir = parent;
  }
  return null;
}

function modeFromFile(p, valid) {
  const c = readJson(p);
  if (c && typeof c.defaultMode === 'string' &&
      valid.includes(c.defaultMode.toLowerCase())) {
    return c.defaultMode.toLowerCase();
  }
  return null;
}

function resolveMode(name) {
  const valid = MODES[name] || [];
  const env = process.env[name.toUpperCase() + '_DEFAULT_MODE'];
  if (env && valid.includes(env.toLowerCase())) return env.toLowerCase();

  const repo = findRepoConfig(process.cwd(), name);
  if (repo) {
    const m = modeFromFile(repo, valid);
    if (m) return m;
  }

  const user = path.join(os.homedir(), '.config', name, 'config.json');
  const m = modeFromFile(user, valid);
  if (m) return m;

  return DEFAULT_MODE;
}

function setMode(name, mode, dir) {
  const valid = MODES[name] || [];
  if (!valid.includes(mode)) {
    process.stderr.write(`[persona] 非法模式 "${mode}"。可用: ${valid.join(' | ')}\n`);
    process.exit(1);
  }
  const target = path.resolve(dir || process.cwd());
  const cfgPath = path.join(target, `.${name}`, 'config.json');
  fs.mkdirSync(path.dirname(cfgPath), { recursive: true });
  fs.writeFileSync(cfgPath, JSON.stringify({ defaultMode: mode }, null, 2) + '\n');
  return cfgPath;
}

// CLI 入口：仅当直接运行本文件时执行（被 require 时不触发）。
if (require.main === module) {
  const [cmd, name, mode, dir] = process.argv.slice(2);

  if (cmd === 'get') {
    console.log(resolveMode(name));
  } else if (cmd === 'set') {
    const cfgPath = setMode(name, mode, dir);
    console.log(`${name} → ${mode}（写入 ${cfgPath}，下次新会话生效）`);
  } else {
    process.stderr.write('[persona] 用法: persona-config.js get|set <name> [mode] [dir]\n');
    process.exit(1);
  }
}

module.exports = { resolveMode, MODES, DEFAULT_MODE };
