/* qbank-core.js — Excel 题库 → 刷题应用题库 的共享转换核心。
 * 浏览器（配合 lib/xlsx.full.min.js）与 Node（tools/build-bank.js）共用同一份逻辑。
 * 用法：QBCore.parse(workbook, XLSX) → {bank:{courses,q}, report}
 *      QBCore.toJsFile(bank, id, label) → data-xxx.js 文本（自带注册进 window.QBANK_EXTRA）
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.QBCore = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var OPT_LETTERS = "ABCDEFGHI";
  var TYPE_KEYS = [["单选", 1], ["多选", 2], ["判断", 3], ["简答", 4]];

  /* 单行文本：压平所有空白 */
  function norm(s) {
    return String(s == null ? "" : s).replace(/\u3000/g, " ").replace(/\s+/g, " ").trim();
  }
  /* 多行文本：保留换行（简答题参考答案用），清理行内多余空白与 \r */
  function normMulti(s) {
    return String(s == null ? "" : s)
      .replace(/\r\n?/g, "\n")
      .replace(/[ \t\u3000]+/g, " ")
      .split("\n").map(function (l) { return l.trim(); }).join("\n")
      .replace(/\n{3,}/g, "\n\n").trim();
  }
  function normHeader(s) {
    return norm(s).replace(/[:：．.\s]/g, "");
  }

  function detectType(text) {
    var t = norm(text);
    if (!t) return 0;
    for (var i = 0; i < TYPE_KEYS.length; i++)
      if (t.indexOf(TYPE_KEYS[i][0]) >= 0) return TYPE_KEYS[i][1];
    if (/问答|主观|简答/.test(t)) return 4;
    if (/^[1234]$/.test(t)) return +t;
    return 0;
  }

  /* 表头 → 列角色。返回 null 表示不是题库表。 */
  function mapHeaders(headerRow) {
    var map = { seq: -1, type: -1, diff: -1, cat: -1, stem: -1, ans: -1, exp: -1, prot: -1, note: -1, opts: {} };
    var found = false;
    for (var c = 0; c < headerRow.length; c++) {
      var h = normHeader(headerRow[c]);
      if (!h) continue;
      if (/^(?:选项)?[A-Ia-i]$/.test(h)) {                  // 选项列：A~I（可带“选项”前缀，不分大小写）
        map.opts[h.replace(/^选项/, "").toUpperCase()] = c; found = true; continue;
      }
      if (h.indexOf("题干") >= 0 || h === "题目" || h === "试题内容") { map.stem = c; found = true; }
      else if (h.indexOf("题型") >= 0) map.type = c;
      else if (h.indexOf("难易") >= 0 || h.indexOf("难度") >= 0) map.diff = c;
      else if (h.indexOf("类别") >= 0 || h.indexOf("课程") >= 0 || h === "知识分类") map.cat = c;
      else if (h.indexOf("答案") >= 0) map.ans = c;
      else if (h.indexOf("出自") >= 0 || h.indexOf("出处") >= 0 || h.indexOf("解析") >= 0 || h.indexOf("依据") >= 0) map.exp = c;
      else if (h.indexOf("保命") >= 0) map.prot = c;
      else if (h.indexOf("序号") >= 0 || h === "编号") map.seq = c;
      else if (h.indexOf("备注") >= 0 || h.indexOf("修编") >= 0 || h.indexOf("说明") >= 0) map.note = c;
    }
    return (found && map.stem >= 0 && map.ans >= 0) ? map : null;
  }

  /* “对/错”类答案归一化；无法识别返回 null */
  function normJudge(text) {
    var t = norm(text).toLowerCase().replace(/[。.\s]/g, "");
    if (/^(对|正确|是|√|✓|t(true)?|y(es)?|a)$/.test(t)) return "对";
    if (/^(错|错误|不对|否|×|x|f(alse)?|n(o)?|b)$/.test(t)) return "错";
    return null;
  }
  /* 从答案单元格里提取选项字母（支持 "ABD"、"A、B、D"、"a b d" 等） */
  function extractLetters(text) {
    var out = [], seen = {};
    var up = String(text == null ? "" : text).toUpperCase();
    for (var i = 0; i < up.length; i++) {
      var ch = up[i];
      if (OPT_LETTERS.indexOf(ch) >= 0 && !seen[ch]) { seen[ch] = 1; out.push(ch); }
    }
    return out.sort();
  }

  /* Excel 日期单元格兜底：极少数情况下 sheet_to_json 可能返回 Date 对象。
   * 主路径用 raw:false，日期单元格已由 SheetJS 按单元格格式输出文本，不会走到这里。 */
  function fmtDate(d) {
    return d.getFullYear() + "年" + (d.getMonth() + 1) + "月" + d.getDate() + "日";
  }
  function cellVal(v) {
    return (v instanceof Date) ? fmtDate(v) : v;
  }

  /* 主入口：wb = SheetJS workbook，XLSX = SheetJS 命名空间 */
  function parse(wb, XLSX) {
    var report = {
      sheets: [], total: 0,
      byType: { 1: 0, 2: 0, 3: 0, 4: 0 },
      warnings: [], errors: []
    };
    var courses = [], q = [];

    wb.SheetNames.forEach(function (sname) {
      var ws = wb.Sheets[sname];
      if (!ws) return;
      /* raw:false：单元格按 Excel 显示格式输出文本（日期单元格因此得到正确文字，
       * 不会泄漏成 44501 这类序列号），文本/数字/公式结果统一为字符串 */
      var grid = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: false });
      if (!grid.length) return;

      /* 定位表头行：包含“题干”的行；找不到则跳过该 sheet（可能是说明页） */
      var hIdx = -1, map = null;
      for (var r = 0; r < Math.min(grid.length, 10); r++) {
        map = mapHeaders(grid[r] || []);
        if (map) { hIdx = r; break; }
      }
      if (hIdx < 0) { report.warnings.push("Sheet「" + sname + "」未找到题库表头，已跳过"); return; }
      var info = { name: sname, rows: 0, kept: 0, skipped: 0 };
      report.sheets.push(info);

      var fallbackType = detectType(sname);   // sheet 名可作题型兜底（“单选题”等）

      for (var r2 = hIdx + 1; r2 < grid.length; r2++) {
        var row = grid[r2] || [];
        var get = function (col) { return col >= 0 ? cellVal(row[col]) : null; };
        var stem = norm(get(map.stem));
        var ansRaw = get(map.ans);
        var hasAny = row.some(function (v) { return v != null && String(v).trim() !== ""; });
        if (!hasAny) continue;                                   // 空行静默跳过
        info.rows++;
        var label = "「" + sname + "」第" + (r2 + 1) + "行";

        if (!stem) { info.skipped++; report.errors.push(label + "：题干为空，已跳过"); continue; }
        var t = detectType(get(map.type)) || fallbackType;
        if (!t) { info.skipped++; report.errors.push(label + "：无法识别题型，已跳过"); continue; }

        var item = { i: 0, c: 0, t: t, s: t === 4 ? normMulti(get(map.stem)) : stem, a: "", e: normMulti(get(map.exp)) };

        /* 选项（单选/多选）：按表头 A~I 顺序收集非空项 */
        var opts = [];
        if (t === 1 || t === 2) {
          var letters = Object.keys(map.opts).sort();
          letters.forEach(function (L) {
            var v = norm(cellVal(row[map.opts[L]]));
            if (v) opts.push(v);
          });
          if (opts.length < 2) { info.skipped++; report.errors.push(label + "：有效选项不足 2 个，已跳过"); continue; }
          item.o = opts;
        }

        /* 答案归一化 */
        if (t === 1 || t === 2) {
          var ls = extractLetters(ansRaw);
          if (!ls.length) { info.skipped++; report.errors.push(label + "：答案中未找到选项字母（原文：" + norm(ansRaw) + "），已跳过"); continue; }
          var maxL = ls[ls.length - 1];
          if (OPT_LETTERS.indexOf(maxL) >= opts.length) {
            info.skipped++; report.errors.push(label + "：答案 " + ls.join("") + " 超出选项数（" + opts.length + " 个），已跳过"); continue;
          }
          if (t === 1 && ls.length > 1) report.warnings.push(label + "：单选题答案含多个字母（" + ls.join("") + "），取第一个");
          if (t === 2 && ls.length < 2) report.warnings.push(label + "：多选题答案只有 1 个字母（" + ls.join("") + "），请人工核对");
          item.a = t === 1 ? ls[0] : ls.join("");
        } else if (t === 3) {
          /* 判断题：答案可能是字母（A=正确/B=错误，见选项列），也可能是文字 */
          var jv = null;
          var ansStr = String(ansRaw == null ? "" : ansRaw).trim();
          if (/^[A-Ia-i]$/.test(ansStr)) {
            var L2 = ansStr.toUpperCase();
            var optText = map.opts[L2] != null ? norm(row[map.opts[L2]]) : "";
            jv = normJudge(optText) || normJudge(ansStr);
          } else {
            jv = normJudge(ansStr);
          }
          if (!jv) { info.skipped++; report.errors.push(label + "：判断题答案无法识别（原文：" + norm(ansRaw) + "），已跳过"); continue; }
          item.a = jv;
        } else {
          item.a = normMulti(ansRaw);
          if (!item.a) { info.skipped++; report.errors.push(label + "：简答题参考答案为空，已跳过"); continue; }
        }

        /* 保命题：是/√/Y/1 → 1 */
        if (map.prot >= 0) {
          var pv = norm(get(map.prot));
          if (/^(是|√|✓|y|yes|true|1)$/i.test(pv)) item.p = 1;
        }

        /* 知识类别 → 课程（按首次出现顺序） */
        var cat = norm(get(map.cat)) || "未分类";
        var ci = courses.indexOf(cat);
        if (ci < 0) { courses.push(cat); ci = courses.length - 1; }
        item.c = ci;

        item.i = q.length + 1;                 // 序号列不管（样例里是公式），统一按行序重排
        q.push(item);
        info.kept++;
        report.byType[t]++;
      }
    });

    report.total = q.length;
    report.courses = courses;
    report.q = q;
    return { bank: { courses: courses, q: q }, report: report };
  }

  /* 生成 data-xxx.js 文本：题库对象 + 自注册到 window.QBANK_EXTRA（index.html 自动收录） */
  function toJsFile(bank, id, label) {
    var safeId = String(id || "bank").toLowerCase().replace(/[^a-z0-9_]/g, "");
    if (!safeId) safeId = "bank";
    var varName = "QBANK_" + safeId.toUpperCase();
    var counts = { 单选: 0, 多选: 0, 判断: 0, 简答: 0 };
    bank.q.forEach(function (x) { counts[{ 1: "单选", 2: "多选", 3: "判断", 4: "简答" }[x.t]]++; });
    var head = "/* " + label + " 题库（由 Excel 题库导入生成，生成时间 " + new Date().toISOString().slice(0, 10) +
      "，共 " + bank.q.length + " 题：单选 " + counts["单选"] + "、多选 " + counts["多选"] +
      "、判断 " + counts["判断"] + "、简答 " + counts["简答"] + "）。请勿手工编辑。 */\n" +
      "/* 使用方法：把本文件放入 data/ 目录，并在 index.html 中 </body> 前加一行：" +
      "<script src=\"data/data-" + safeId + ".js\"></" + "script> 即可自动出现在题库列表。 */\n";
    var body = "window." + varName + " = " + JSON.stringify(Object.assign({ id: safeId, name: label }, bank)) + ";\n" +
      "(function(){ window.QBANK_EXTRA = window.QBANK_EXTRA || [];\n" +
      "  window.QBANK_EXTRA.push({ id: \"" + safeId + "\", label: \"" + String(label).replace(/["\\]/g, "") + "\", raw: window." + varName + " });\n" +
      "})();\n";
    return head + body;
  }

  return {
    parse: parse,
    toJsFile: toJsFile,
    detectType: detectType,
    normJudge: normJudge,
    extractLetters: extractLetters,
    mapHeaders: mapHeaders
  };
});
