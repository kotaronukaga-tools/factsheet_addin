/* ファクトシート整形アドイン
 * 変換ルールの原本は factsheet-tool/RULES.md および prototype/transform.py。
 * このファイルはそのロジックの Office.js 移植版。
 */
"use strict";

const BODY_FONT = "Univers 55";
const HEADER_FONT = "Univers";
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

Office.onReady((info) => {
  const btn = document.getElementById("run");
  if (info.host === Office.HostType.Word) {
    btn.disabled = false;
    btn.addEventListener("click", run);
  } else {
    setStatus('<div class="error">このアドインはWord専用です。</div>');
  }
});

function setStatus(html) {
  document.getElementById("status").innerHTML = html;
}

function esc(s) {
  return String(s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}

/* ---------- 抽出ロジック (transform.py と同一ルール) ---------- */

function splitLines(text) {
  // Wordの段落内改行(ソフト改行)は \u000b、環境により \r \n も現れる
  return text.split(/[\u000b\r\n\u2028]+/).map(s => s.trim()).filter(Boolean);
}

const TITLE_YEAR_RE = /^(.*?),\s*((?:ca\.\s*)?\d{4}(?:\s*[-–]\s*\d{2,4})?)$/;
const CODE_RE = /^[（(]?[A-Z]{2,4}_\d+[a-z]?[）)]?$/;
const SIZE_RE = /\d+(\.\d+)?\s*(x|×)\s*\d+/;
const PRICE_RE = /JPY|USD|\$|¥|円/;

function extractFields(detailText, priceText) {
  const lines = splitLines(detailText);
  const fields = { artist: null, title: null, year: null, technique: null, size: null, code: null };
  const rest = [];
  for (const txt of lines) {
    if (fields.artist === null) {
      fields.artist = txt; // Artlogic書き出しは1行目が作家名(太字)
      continue;
    }
    const m = txt.match(TITLE_YEAR_RE);
    if (fields.title === null && m) {
      fields.title = m[1].trim();
      fields.year = m[2];
      continue;
    }
    rest.push(txt);
  }
  for (const txt of rest) {
    if (fields.code === null && CODE_RE.test(txt)) {
      fields.code = txt.replace(/[（()）]/g, "");
    } else if (fields.size === null && SIZE_RE.test(txt)) {
      fields.size = txt;
    } else if (fields.technique === null) {
      fields.technique = txt;
    }
  }
  fields.priceRaw = (priceText || "").trim();
  return fields;
}

/* 価格整形。併記(JPY+USD)は書式未確定のため無変換で警告 */
function normalizePrice(text) {
  const t = (text || "").trim();
  if (!t) return { price: t, warn: null };
  const hasJpy = /JPY|¥|円/.test(t);
  const hasUsd = /USD|\$/.test(t);
  if (hasJpy && hasUsd) {
    return { price: t, warn: "円・USD併記の価格は書式が未確定のため変更していません。手動で確認してください。" };
  }
  let p = t.replace(/USD\s*\$\s*/, "USD ").replace(/^\$\s*/, "USD ");
  if (!/excl\.?\s*TAX/i.test(p)) {
    p += " excl. TAX";
  } else {
    p = p.replace(/\s*excl\.?\s*TAX\.?/i, " excl. TAX");
  }
  return { price: p, warn: null };
}

/* ---------- OOXML生成 (Editテンプレートと同一構造) ---------- */

function fontXml(font) {
  return `<w:rFonts w:ascii="${font}" w:hAnsi="${font}"/>`;
}

function runXml(text, opts = {}) {
  const italic = opts.italic ? "<w:i/>" : "";
  const sz = opts.sz ? `<w:sz w:val="${opts.sz}"/><w:szCs w:val="${opts.sz}"/>` : "";
  return `<w:r><w:rPr>${fontXml(BODY_FONT)}${italic}${sz}</w:rPr>` +
    `<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
}

function brRunXml() {
  return `<w:r><w:rPr>${fontXml(BODY_FONT)}</w:rPr><w:br/></w:r>`;
}

function paraXml(runsXml, jc) {
  const jcXml = jc ? `<w:jc w:val="${jc}"/>` : "";
  return `<w:p><w:pPr><w:contextualSpacing/>${jcXml}` +
    `<w:rPr>${fontXml(BODY_FONT)}</w:rPr></w:pPr>${runsXml}</w:p>`;
}

function buildBodyOoxml(fields, price) {
  const titlePara = paraXml(
    brRunXml() + runXml(fields.title || "", { italic: true, sz: 32 }), "left");
  const detailPara = paraXml(
    runXml(fields.year || "") + brRunXml() +
    runXml(fields.technique || "") + brRunXml() +
    runXml(fields.size || "") + brRunXml() +
    runXml(fields.code ? `（${fields.code}）` : ""), "left");
  const emptyPara = paraXml("");
  const pricePara = paraXml(runXml(price) + brRunXml(), "left");
  const body = titlePara + detailPara + emptyPara + pricePara + emptyPara;
  return (
    '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
    '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">' +
    '<pkg:xmlData>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/word/document.xml"/>' +
    '</Relationships></pkg:xmlData></pkg:part>' +
    '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">' +
    `<pkg:xmlData><w:document xmlns:w="${W_NS}"><w:body>${body}</w:body></w:document></pkg:xmlData>` +
    '</pkg:part></pkg:package>'
  );
}

/* ---------- メイン処理 ---------- */

async function run() {
  const btn = document.getElementById("run");
  btn.disabled = true;
  setStatus("整形中…");
  try {
    await Word.run(async (ctx) => {
      const paras = ctx.document.body.paragraphs;
      paras.load("items/text");
      await ctx.sync();

      const items = paras.items;
      let imagePara = null, detailPara = null, pricePara = null;
      const others = [];
      for (const p of items) {
        const t = p.text.trim();
        if (!t) {
          if (imagePara === null && detailPara === null) imagePara = p;
          else others.push(p);
        } else if (PRICE_RE.test(t) && detailPara !== null) {
          if (pricePara === null) pricePara = p; else others.push(p);
        } else if (detailPara === null) {
          detailPara = p;
        } else {
          others.push(p);
        }
      }
      if (!detailPara) throw new Error("作品情報の段落が見つかりません。対象のファクトシートを開いた状態で実行してください。");

      // 整形済みファイルの再実行ガード: 1行目が年のみなら整形済みとみなす
      const firstLine = splitLines(detailPara.text)[0] || "";
      if (/^(?:ca\.\s*)?\d{4}(?:\s*[-–]\s*\d{2,4})?$/.test(firstLine)) {
        throw new Error("このファイルは既に整形済みのようです（本文が年から始まっています）。");
      }

      const fields = extractFields(detailPara.text, pricePara ? pricePara.text : "");
      const missing = ["artist", "title", "year", "technique", "size", "code"]
        .filter(k => !fields[k]);
      const { price, warn } = normalizePrice(fields.priceRaw);
      const warns = [];
      if (warn) warns.push(warn);
      if (missing.length) {
        const ja = { artist: "作家名", title: "作品名", year: "年", technique: "技法", size: "サイズ", code: "管理番号" };
        warns.push("抽出できなかった項目: " + missing.map(k => ja[k]).join("・") + "（該当箇所が空欄になります。手動で補ってください）");
      }

      // 本文再構成: 画像段落以外を削除し、テンプレート構造のOOXMLを挿入
      const ooxml = buildBodyOoxml(fields, price);
      if (imagePara) {
        imagePara.insertOoxml(ooxml, Word.InsertLocation.after);
      } else {
        ctx.document.body.insertOoxml(ooxml, Word.InsertLocation.start);
      }
      detailPara.delete();
      if (pricePara) pricePara.delete();
      for (const p of others) p.delete();

      // 本文フォント統一(画像段落の改行runにも適用)。ヘッダー/フッターには影響しない
      ctx.document.body.font.name = BODY_FONT;

      // ヘッダーの作家名差し替え(書式・罫線・空段落は保持)
      let headerDone = false;
      if (fields.artist) {
        const header = ctx.document.sections.getFirst()
          .getHeader(Word.HeaderFooterType.primary);
        const hParas = header.paragraphs;
        hParas.load("items/text");
        await ctx.sync();
        for (const hp of hParas.items) {
          if (hp.text.trim()) {
            const r = hp.getRange(Word.RangeLocation.whole);
            r.insertText(fields.artist, Word.InsertLocation.replace);
            hp.font.name = HEADER_FONT;
            hp.font.size = 22;
            hp.font.bold = false;
            headerDone = true;
            break;
          }
        }
      }
      if (!headerDone) {
        warns.push("ヘッダー内に差し替え対象の作家名が見つかりませんでした。手動で確認してください。");
      }
      await ctx.sync();

      // 結果表示
      const ja = [["作家名", fields.artist], ["作品名", fields.title], ["年", fields.year],
        ["技法", fields.technique], ["サイズ", fields.size], ["管理番号", fields.code],
        ["価格", price]];
      let html = '<div class="ok">✓ 整形が完了しました。内容を確認してから保存してください。</div>';
      html += '<table class="fields">' + ja.map(([k, v]) =>
        `<tr><td>${k}</td><td>${esc(v || "（空欄）")}</td></tr>`).join("") + "</table>";
      for (const wmsg of warns) html += `<div class="warn">⚠ ${esc(wmsg)}</div>`;
      setStatus(html);
    });
  } catch (e) {
    setStatus(`<div class="error">エラー: ${esc(e.message || e)}</div>` +
      '<div class="note">解決しない場合はこのファイルをPR担当（Yui）に共有してください。</div>');
  } finally {
    btn.disabled = false;
  }
}
