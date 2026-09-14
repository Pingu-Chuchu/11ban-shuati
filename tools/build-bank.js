#!/usr/bin/env node
/* build-bank.js — 把 Excel 题库转换为应用题库文件（与网页内导入共用 lib/qbank-core.js）。
 *
 * 用法：
 *   node tools/build-bank.js <Excel文件> <输出文件> <题库短id> <题库名称>
 * 例：
 *   node tools/build-bank.js "题库（2026年版）.xlsx" data/data-ag.js ag 安规
 *
 * 短 id 规则：小写字母/数字/下划线，用于 URL（#ag）和进度存储隔离。
 */
"use strict";
const path = require("path");
const fs = require("fs");

const XLSX = require(path.join(__dirname, "..", "lib", "xlsx.full.min.js"));
const QBCore = require(path.join(__dirname, "..", "lib", "qbank-core.js"));

const [, , inExcel, outFile, id, label] = process.argv;
if (!inExcel || !outFile || !id || !label) {
  console.error("用法: node tools/build-bank.js <Excel文件> <输出文件.js> <短id> <题库名称>");
  process.exit(2);
}
if (!/^[a-z0-9_]+$/.test(id)) {
  console.error("短 id 只能是小写字母/数字/下划线，收到: " + id);
  process.exit(2);
}

const wb = XLSX.read(fs.readFileSync(inExcel), { type: "buffer" });
const { bank, report } = QBCore.parse(wb, XLSX);

/* 报告写到 UTF-8 文件，避免 Windows 控制台中文乱码 */
const lines = [];
lines.push(`Sheet 解析: ` + report.sheets.map(s => `${s.name}(读 ${s.rows} 行, 收 ${s.kept}, 跳过 ${s.skipped})`).join(" | "));
lines.push(`共 ${report.total} 题：单选 ${report.byType[1]}、多选 ${report.byType[2]}、判断 ${report.byType[3]}、简答 ${report.byType[4]}`);
lines.push(`课程(${report.courses.length}): ${report.courses.join(" | ")}`);
if (report.warnings.length) {
  lines.push(`警告 ${report.warnings.length} 条:`);
  report.warnings.slice(0, 50).forEach(w => lines.push("  ⚠ " + w));
}
if (report.errors.length) {
  lines.push(`错误 ${report.errors.length} 条:`);
  report.errors.slice(0, 50).forEach(e => lines.push("  ✗ " + e));
}
lines.push("");
if (!report.total) {
  lines.push("没有解析到任何题目，未生成文件。");
  fs.writeFileSync(outFile.replace(/\.js$/, "") + ".report.txt", lines.join("\n"), "utf8");
  console.error(lines.join("\n"));
  process.exit(1);
}

fs.mkdirSync(path.dirname(path.resolve(outFile)), { recursive: true });
fs.writeFileSync(outFile, QBCore.toJsFile(bank, id, label), "utf8");
lines.push(`已生成: ${outFile}`);
fs.writeFileSync(outFile.replace(/\.js$/, "") + ".report.txt", lines.join("\n"), "utf8");
console.log("OK, report -> " + outFile.replace(/\.js$/, "") + ".report.txt");
