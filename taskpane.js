/* ファクトシート整形アドイン
 * 変換ルールの原本は factsheet-tool/RULES.md および prototype/transform.py。
 * このファイルはそのロジックの Office.js 移植版。
 */
"use strict";

const VERSION = "1.1.7";
const BODY_FONT = "Univers 55";
const HEADER_FONT = "Univers";

// バージョン表示はOffice初期化を待たず即時に(古いキャッシュ判別のため)
(function showVersion() {
  const el = document.getElementById("ver");
  if (el) el.textContent = "v" + VERSION;
})();
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PRICE_TAB_POS = "5400"; // 額装併記の2列整列タブストップ(twips)
// 価格が読み取れない場合のプレースホルダー(最終書式に合わせ、そのまま数字を上書きできる形に)
const AMT_PLACEHOLDER = "0,000";   // 金額
const CUR_PLACEHOLDER = "XXX";     // 通貨コード

Office.onReady((info) => {
  console.log(`[ファクトシート整形] version ${VERSION} loaded`);
  const btn = document.getElementById("run");
  if (info.host === Office.HostType.Word) {
    btn.disabled = false;
    btn.addEventListener("click", run);
    // Word Online(ブラウザ版)は埋め込み画像(VML)を保持できず脱落するため警告
    try {
      if (Office.context.platform === Office.PlatformType.OfficeOnline) {
        const el = document.getElementById("onlineWarn");
        if (el) el.style.display = "block";
      }
    } catch (e) { /* platform判定不可時は無視 */ }
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
// 価格段落の判定に使う(通貨コード or 記号 or 「円」)。GBP/EUR対応で拡張
const PRICE_RE = /JPY|USD|GBP|EUR|[$¥£€]|円/;

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

// 通貨コードと記号。JPYは国内通貨、それ以外(USD/GBP/EUR)は外貨として扱う。
// 併記(dual)は「JPY + 外貨1種」を想定。外貨の追加はこの表とFOREIGN_CURRENCIESに足すだけ。
const CURRENCY_SYMBOLS = { USD: "\\$", GBP: "£", EUR: "€", JPY: "¥" };
const FOREIGN_CURRENCIES = ["USD", "GBP", "EUR"];

function parsePriceComponents(text) {
  const t = text || "";
  const base = {}, framing = {};
  for (const code of Object.keys(CURRENCY_SYMBOLS)) {
    const sym = CURRENCY_SYMBOLS[code];
    // Framing額を先に抜き出し、base検索がFraming額を拾わないようその通貨のFramingを除去
    const mF = t.match(new RegExp("Framing\\s+" + code + "\\s*(?:" + sym + ")?\\s*" + NUM, "i"));
    const rest = t.replace(
      new RegExp("\\+?\\s*Framing\\s+" + code + "\\s*(?:" + sym + ")?\\s*" + NUM, "ig"), "");
    const mB = rest.match(new RegExp(code + "\\s*(?:" + sym + ")?\\s*" + NUM, "i"));
    base[code] = mB ? mB[1] : null;
    framing[code] = mF ? mF[1] : null;
  }
  return { base, framing };
}

function detectPriceMode(text) {
  const t = text || "";
  const hasJpy = /JPY|¥|円/.test(t);
  let foreign = null;
  for (const code of FOREIGN_CURRENCIES) {
    if (new RegExp(code + "|" + CURRENCY_SYMBOLS[code]).test(t)) { foreign = code; break; }
  }
  return { dual: hasJpy && foreign !== null, framing: /Framing/i.test(t), foreign };
}

/* モードに応じた価格行を組み立てる。各行は {type:"line",text} または {type:"cols",cols:[左,右]} */
function buildPriceLines(comp, mode) {
  const warns = [];
  const A = AMT_PLACEHOLDER;
  const base = comp.base, framing = comp.framing;
  const foreign = mode.foreign;
  const jpy = base.JPY;
  const fx = foreign ? base[foreign] : null;
  if (!mode.framing) {
    if (mode.dual) {
      const j = jpy || A, u = fx || A;
      if (!jpy) warns.push("JPY金額が元ファイルに無いため 0,000 にしました。手動で入力してください。");
      if (!fx) warns.push(`${foreign}金額が抽出できませんでした。手動で確認してください。`);
      return { lines: [{ type: "line", text: `JPY ${j} / ${foreign} ${u} excl. TAX` }], warns };
    }
    if (fx) return { lines: [{ type: "line", text: `${foreign} ${fx} excl. TAX` }], warns };
    if (jpy) return { lines: [{ type: "line", text: `JPY ${jpy} excl. TAX` }], warns };
    // 価格が全く読み取れない場合はテンプレート文言を挿入(通貨・金額とも要上書き)
    warns.push("価格が読み取れませんでした。通貨・金額を手動で入力してください。");
    return { lines: [{ type: "line", text: `${CUR_PLACEHOLDER} ${A} excl. TAX` }], warns };
  }
  // 額装あり
  if (mode.dual) {
    const ub = fx || A, jb = jpy || A;
    const uf = (foreign ? framing[foreign] : null) || A, jf = framing.JPY || A;
    if (!(jpy && framing.JPY)) warns.push("JPY金額が不足しています。手動で入力してください。");
    if (!(foreign && framing[foreign])) warns.push(`${foreign} Framing金額が抽出できませんでした。手動で確認してください。`);
    return {
      lines: [
        { type: "cols", cols: [`${foreign} ${ub}`, `JPY ${jb}`] },
        { type: "cols", cols: [`+ Framing ${foreign} ${uf} excl. TAX`, `+ Framing JPY ${jf} excl. TAX`] },
      ], warns,
    };
  }
  const cur = fx ? foreign : (jpy ? "JPY" : CUR_PLACEHOLDER);
  const b = fx || jpy || A;
  const fr = (foreign ? framing[foreign] : null) || framing.JPY || A;
  return {
    lines: [
      { type: "cols", cols: [`${cur} ${b}`] },
      { type: "cols", cols: [`+ Framing ${cur} ${fr} excl. TAX`] },
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

function paraXml(runsXml, jc, tabPos, sz) {
  const tabs = tabPos ? `<w:tabs><w:tab w:val="left" w:pos="${tabPos}"/></w:tabs>` : "";
  const jcXml = jc ? `<w:jc w:val="${jc}"/>` : "";
  const szXml = sz ? `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>` : "";
  return `<w:p><w:pPr>${tabs}<w:contextualSpacing/>${jcXml}` +
    `<w:rPr>${fontXml(BODY_FONT)}${szXml}</w:rPr></w:pPr>${runsXml}</w:p>`;
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
  // タイトルと制作年の間の1行空け(10.5pt = sz21)
  const titleGapPara = paraXml("", "left", null, 21);
  let detailRuns = runXml(fields.year || "");
  for (const part of fields.middle) detailRuns += brRunXml() + runXml(part);
  detailRuns += brRunXml() + runXml(fields.code ? `（${fields.code}）` : "");
  const detailPara = paraXml(detailRuns, "left");
  const emptyPara = paraXml("");
  const priceParas = priceParasXml(priceLines);
  const body = titlePara + titleGapPara + detailPara + emptyPara + priceParas + emptyPara;
  return (
    '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
    '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">' +
    '<pkg:xmlData>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships></pkg:xmlData></pkg:part>' +
    '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">' +
    `<pkg:xmlData><w:document xmlns:w="${W_NS}"><w:body>${body}</w:body></w:document></pkg:xmlData>` +
    '</pkg:part></pkg:package>'
  );
}

/* ---------- メイン処理 ---------- */

/* エラー詳細をパネルとコンソールの両方に出す(Word Onlineでのデバッグ用) */
function describeError(e, stage) {
  console.error(`[ファクトシート整形] エラー stage=${stage}`, e);
  let html = `<div class="error">エラー（${esc(stage)}）: ${esc(e.message || e)}</div>`;
  if (typeof OfficeExtension !== "undefined" && e instanceof OfficeExtension.Error) {
    const d = e.debugInfo || {};
    const details = {
      code: e.code,
      errorLocation: d.errorLocation,
      statement: d.statement,
      surroundingStatements: d.surroundingStatements,
    };
    console.error("[ファクトシート整形] debugInfo", d);
    html += `<pre class="debug">${esc(JSON.stringify(details, null, 2))}</pre>`;
    html += '<div class="note">上記の詳細（特にerrorLocation）をコピーしてKawasakiに共有してください。</div>';
  } else {
    html += '<div class="note">解決しない場合はこのファイルをKawasakiに共有してください。</div>';
  }
  return html;
}

/* insertOoxmlのホスト差異を吸収するフォールバック。旧本文段落を削除した後に呼ぶ前提で、
 * いずれの方式も「画像段落の直後(=末尾)」に挿入される。
 * Body.insertOoxml(End) や Range.insertOoxml は Paragraph.insertOoxml(After) より広くサポートされる。 */
async function insertOoxmlWithFallback(ctx, ooxml, imagePara) {
  const attempts = [
    () => ctx.document.body.insertOoxml(ooxml, Word.InsertLocation.end),
  ];
  if (imagePara) {
    attempts.push(() =>
      imagePara.getRange(Word.RangeLocation.after).insertOoxml(ooxml, Word.InsertLocation.after));
  }
  let lastErr;
  for (let i = 0; i < attempts.length; i++) {
    try {
      attempts[i]();
      await ctx.sync();
      console.log(`[ファクトシート整形] insertOoxml 方式${i + 1} 成功`);
      return;
    } catch (e) {
      lastErr = e;
      console.warn(`[ファクトシート整形] insertOoxml 方式${i + 1} 失敗:`, e && e.message);
    }
  }
  throw lastErr;
}

async function run() {
  const btn = document.getElementById("run");
  btn.disabled = true;
  setStatus("整形中…");
  let stage = "初期化";
  try {
    await Word.run(async (ctx) => {
      stage = "本文の読み取り";
      const paras = ctx.document.body.paragraphs;
      paras.load("items/text");
      await ctx.sync();
      stage = "項目の抽出";

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

      // 価格の形式は常に元ファイルから自動判定する(併記/額装/外貨コード)
      const comp = parsePriceComponents(fields.priceRaw);
      const mode = detectPriceMode(fields.priceRaw);
      const { lines: priceLines, warns } = buildPriceLines(comp, mode);

      if (missing.length) {
        const ja = { artist: "作家名", title: "作品名", year: "年", code: "管理番号", "technique/size": "技法・サイズ" };
        warns.push("抽出できなかった項目: " + missing.map(k => ja[k]).join("・") + "（該当箇所が空欄になります。手動で補ってください）");
      }

      // 旧本文段落を先に削除(必要データは抽出済み)。画像段落だけ残す
      stage = "旧本文段落の削除";
      detailPara.delete();
      if (pricePara) pricePara.delete();
      for (const p of others) p.delete();
      await ctx.sync();

      // テンプレート構造のOOXMLを画像の直後(=末尾)に挿入。
      // Paragraph.insertOoxml(After) はMac/Onlineで拒否されるため方式を切り替え+フォールバック
      stage = "本文の再構成(OOXML挿入)";
      const ooxml = buildBodyOoxml(fields, priceLines);
      console.log("[ファクトシート整形] fields:", JSON.stringify(fields));
      console.log("[ファクトシート整形] ooxml:", ooxml);
      await insertOoxmlWithFallback(ctx, ooxml, imagePara);
      // 挿入する各段落・runには既にUnivers 55を付与済みのため、
      // body全体へのフォント適用は行わない(画像段落に触れるとWord OnlineがVML画像を脱落させるため)

      // ヘッダーの作家名差し替え(書式・罫線・空段落は保持)
      stage = "ヘッダーの作家名差し替え";
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
      stage = "結果表示";

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
    setStatus(describeError(e, stage));
  } finally {
    btn.disabled = false;
  }
}
