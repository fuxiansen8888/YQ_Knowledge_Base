#!/usr/bin/env node
// activate-personas — ZCode SessionStart hook。
//
// 每次新会话（startup/resume/clear/compact）触发：按当前模式读取
// .zcode/personas/ 下的规则文件，把生效模式的规则作为 additionalContext
// 注入对话。模式解析见 persona-config.js。
//
// 输出 JSON（hook schema 严格，仅此一个 key）。任何异常都静默降级为 {}，
// 保证会话启动永不被阻塞。

const fs = require('fs');
const path = require('path');
const { resolveMode } = require('./persona-config');

const RULES = {
  caveman: path.join(__dirname, '..', 'personas', 'caveman-rules.md'),
  ponytail: path.join(__dirname, '..', 'personas', 'ponytail-rules.md'),
};

// 规则文件按当前强度过滤：强度表格只留当前级别行，示例块只留当前级别行。
function filterRules(file, mode) {
  let body;
  try {
    body = fs.readFileSync(file, 'utf8');
  } catch (e) {
    return '';
  }
  return body.split('\n').reduce((acc, line) => {
    const tableRow = line.match(/^\|\s*\*\*(\S+?)\*\*\s*\|/);
    if (tableRow) {
      if (tableRow[1] === mode) acc.push(line);
      return acc;
    }
    const exampleRow = line.match(/^- (\S+?):\s/);
    if (exampleRow) {
      if (exampleRow[1] === mode) acc.push(line);
      return acc;
    }
    acc.push(line);
    return acc;
  }, []).join('\n');
}

let output = {};
try {
  const sections = [];
  const cavemanMode = resolveMode('caveman');
  if (cavemanMode !== 'off') {
    sections.push(`CAVEMAN MODE ACTIVE — level: ${cavemanMode}\n\n` +
      filterRules(RULES.caveman, cavemanMode));
  }
  const ponytailMode = resolveMode('ponytail');
  if (ponytailMode !== 'off') {
    sections.push(`PONYTAIL ACTIVE — level: ${ponytailMode}\n\n` +
      filterRules(RULES.ponytail, ponytailMode));
  }
  if (sections.length > 0) {
    output = { additionalContext: sections.join('\n\n---\n\n') };
  }
} catch (e) {
  // 静默降级：hook 失败不阻塞会话启动
}

process.stdout.write(JSON.stringify(output));
