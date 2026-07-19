/* ファクトシート整形アドイン
 * 変換ルールの原本は factsheet-tool/RULES.md および prototype/transform.py。
 * このファイルはそのロジックの Office.js 移植版。
 */
"use strict";

const BODY_FONT = "Univers 55";
const HEADER_FONT = "Univers";
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PRICE_TAB_POS = "5400"; // 額装併記の2列整列タブストップ(twips)
const PLACEHOLDER = "（要入力）";

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
  // Wordの段落内ソフト改行はU+000B。CR/LF/U+2028も区切りに含める
  return text.split(/[\u000b\r\n\u2028]+/).map(s => s.trim()).filter(Boolean);
}

function cleanLine(txt) {
  // 軽微な体裁補正: 読点・カンマ前の余分な空白除去、連続空白の圧縮
  return txt.replace(/\s+([,，、.。])/g, "$1").replace(/[ 　]{2,}/g, " ").trim();
}

const TITLE_YEAR_RE = /^(.*?),\s*((?:ca\.\s*)?\d{4}(?:\s*[-–]\s*\d{2,4})?)$/;
const CODE_RE = /^[（(]?[A-Z]{2,4}_\d+[a-z]?[）)]?$/;
const PRICE_RE = /JPY|USD|\$|¥|円/;

function extractFields(detailText, priceText) {
  const lines = splitLines(detailText);
  const fields = { artist: null, title: null, year: null, code: null, middle: [] };
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
  // 管理番号(コード)を分離。それ以外(技法・サイズ等)は順序保持で中間行に
  for (const txt of rest) {
    if (fields.code === null && CODE_RE.test(txt)) {
      fields.code = txt.replace(/[（()）]/g, "");
    } else {
      fields.middle.push(cleanLine(txt));
    }
  }
  fields.priceRaw = (priceText || "").trim();
  return fields;
}

/* ---------- 価格ロジック ---------- */

const NUM = "([\\d,]+(?:\\.\\d+)?)";

function parsePriceComponents(text) {
  const t = text || "";
  const mUf = t.match(new RegExp("Framing\\s+USD\\s*\\$?\\s*" + NUM, "i"));
  const mJf = t.match(new RegExp("Framing\\s+JPY\\s*¥?\\s*" + NUM, "i"));
  const rest = t
    .replace(new RegExp("\\+?\\s*Framing\\s+USD\\s*\\$?\\s*" + NUM, "ig"), "")
    .replace(new RegExp("\\+?\\s*Framing\\s+JPY\\s*¥?\\s*" + NUM, "ig"), "");
  const mUb = rest.match(new RegExp("USD\\s*\\$?\\s*" + NUM, "i"));
  const mJb = rest.match(new RegExp("JPY\\s*¥?\\s*" + NUM, "i"));
  return {
    usdBase: mUb ? mUb[1] : null,
    usdFraming: mUf ? mUf[1] : null,
    jpyBase: mJb ? mJb[1] : null,
    jpyFraming: mJf ? mJf[1] : null,
  };
}

function detectPriceMode(text) {
  const t = text || "";
  return { dual: /USD|\$/.test(t) && /JPY|¥|円/.test(t), framing: /Framing/i.test(t) };
}

/* モードに応じた価格行を組み立てる。各行は {type:"line",text} または {type:"cols",cols:[左,右]} */
function buildPriceLines(comp, dual, framing) {
  const warns = [];
  const P = PLACEHOLDER;
  if (!framing) {
    if (dual) {
      const j = comp.jpyBase || P, u = comp.usdBase || P;
      if (!comp.jpyBase) warns.push("JPY金額が元ファイルに無いため（要入力）にしました。手動で入力してください。");
      if (!comp.usdBase) warns.push("USD金額が抽出できませんでした。手動で確認してください。");
      return { lines: [{ type: "line", text: `JPY ${j} / USD ${u} excl. TAX` }], warns };
    }
    if (comp.usdBase) return { lines: [{ type: "line", text: `USD ${comp.usdBase} excl. TAX` }], warns };
    if (comp.jpyBase) return { lines: [{ type: "line", text: `JPY ${comp.jpyBase} excl. TAX` }], warns };
    warns.push("価格が抽出できませんでした。手動で入力してください。");
    return { lines: [{ type: "line", text: P }], warns };
  }
  // 額装あり
  if (dual) {
    const ub = comp.usdBase || P, jb = comp.jpyBase || P;
    const uf = comp.usdFraming || P, jf = comp.jpyFraming || P;
    if (!(comp.jpyBase && comp.jpyFraming)) warns.push("JPY金額が不足しています。手動で入力してください。");
    if (!comp.usdFraming) warns.push("USD Framing金額が抽出できませんでした。手動で確認してください。");
    return {
      lines: [
        { type: "cols", cols: [`USD ${ub}`, `JPY ${jb}`] },
        { type: "cols", cols: [`+ Framing USD ${uf} excl. TAX`, `+ Framing JPY ${jf} excl. TAX`] },
      ], warns,
    };
  }
  const cur = comp.usdBase ? "USD" : (comp.jpyBase ? "JPY" : "USD");
  const base = comp.usdBase || comp.jpyBase || P;
  const fram = comp.usdFraming || comp.jpyFraming || P;
  return {
    lines: [
      { type: "cols", cols: [`${cur} ${base}`] },
      { type: "cols", cols: [`+ Framing ${cur} ${fram} excl. TAX`] },
    ], warns,
  };
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

function tabRunXml() {
  return `<w:r><w:rPr>${fontXml(BODY_FONT)}</w:rPr><w:tab/></w:r>`;
}

function paraXml(runsXml, jc, tabPos) {
  const tabs = tabPos ? `<w:tabs><w:tab w:val="left" w:pos="${tabPos}"/></w:tabs>` : "";
  const jcXml = jc ? `<w:jc w:val="${jc}"/>` : "";
  return `<w:p><w:pPr>${tabs}<w:contextualSpacing/>${jcXml}` +
    `<w:rPr>${fontXml(BODY_FONT)}</w:rPr></w:pPr>${runsXml}</w:p>`;
}

function priceParasXml(priceLines) {
  let out = "";
  priceLines.forEach((ln, i) => {
    const last = i === priceLines.length - 1;
    if (ln.type === "line") {
      out += paraXml(runXml(ln.text) + (last ? brRunXml() : ""), "left");
    } else {
      const cols = ln.cols;
      let runs = runXml(cols[0]);
      if (cols.length > 1) runs += tabRunXml() + runXml(cols[1]);
      if (last) runs += brRunXml();
      out += paraXml(runs, "left", cols.length > 1 ? PRICE_TAB_POS : null);
    }
  });
  return out;
}

function buildBodyOoxml(fields, priceLines) {
  const titlePara = paraXml(
    brRunXml() + runXml(fields.title || "", { italic: true, sz: 32 }), "left");
  let detailRuns = runXml(fields.year || "");
  for (const part of fields.middle) detailRuns += brRunXml() + runXml(part);
  detailRuns += brRunXml() + runXml(fields.code ? `（${fields.code}）` : "");
  const detailPara = paraXml(detailRuns, "left");
  const emptyPara = paraXml("");
  const priceParas = priceParasXml(priceLines);
  const body = titlePara + detailPara + emptyPara + priceParas + emptyPara;
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

function resolveMode(id, detected, values) {
  const v = document.getElementById(id) ? document.getElementById(id).value : "auto";
  return v === "auto" ? detected : values[v];
}

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
      const missing = ["artist", "title", "year", "code"].filter(k => !fields[k]);
      if (!fields.middle.length) missing.push("technique/size");

      // 価格モード: UIで「自動判定」なら元ファイルから判定、明示指定ならそれを使う
      const comp = parsePriceComponents(fields.priceRaw);
      const detected = detectPriceMode(fields.priceRaw);
      // UIは2択: 「単一通貨/なし」(value=auto)は自動判定、「併記/あり」は強制指定
      const dual = resolveMode("currencyMode", detected.dual, { dual: true });
      const framing = resolveMode("framingMode", detected.framing, { framed: true });
      const { lines: priceLines, warns } = buildPriceLines(comp, dual, framing);

      if (missing.length) {
        const ja = { artist: "作家名", title: "作品名", year: "年", code: "管理番号", "technique/size": "技法・サイズ" };
        warns.push("抽出できなかった項目: " + missing.map(k => ja[k]).join("・") + "（該当箇所が空欄になります。手動で補ってください）");
      }

      // 本文再構成: 画像段落以外を削除し、テンプレート構造のOOXMLを挿入
      const ooxml = buildBodyOoxml(fields, priceLines);
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
      const priceText = priceLines.map(ln =>
        ln.type === "line" ? ln.text : ln.cols.join("　　")).join(" / ");
      const rows = [["作家名", fields.artist], ["作品名", fields.title], ["年", fields.year],
        ["技法・サイズ", fields.middle.join(" / ")], ["管理番号", fields.code],
        ["価格", priceText]];
      let html = '<div class="ok">✓ 整形が完了しました。内容を確認してから保存してください。</div>';
      html += '<table class="fields">' + rows.map(([k, v]) =>
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
