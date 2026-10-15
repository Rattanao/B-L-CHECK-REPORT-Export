/* B-L-CHECK — ตรรกะตรวจ B/L ของลูกค้า (CLIENT) เทียบ DRAFT ของ HAL ครบ 27 หัวข้อ + RIDER / AS PER ATTACHED
   พอร์ตจาก scripts/bl_checker.py ให้ทำงานในเบราว์เซอร์ล้วน ๆ (pdf.js) ไม่มีการอัปโหลดไฟล์ไปที่ใด
   ถ้าแก้ตรรกะ ให้แก้ทั้งสองไฟล์ให้ตรงกัน */
(function (root) {
  "use strict";

  // ───────────── 27 หัวข้อตาม BL DETAILS.xlsx ─────────────
  const FIELDS = [
    ["bl_no", "B/L NO."],
    ["shipper", "SHIPPER"],
    ["consignee", "CONSIGNEE"],
    ["notify", "NOTIFY PARTY"],
    ["pre_carriage", "PRE CARRIAGE BY"],
    ["place_receipt", "PLACE OF RECEIPT"],
    ["vessel_voy", "OCEAN VESSEL VOY NO."],
    ["pol", "PORT OF LOADING"],
    ["pod", "PORT OF DISCHARGE"],
    ["place_delivery", "PLACE OF DELIVERY"],
    ["final_dest", "FINAL DESTINATION (FOR THE MERCHANTS REFERENCE ONLY)"],
    ["container_seal", "CONTAINER NO. / SEAL NO."],
    ["no_pkgs", "NO. OF CONTAINERS OR PKGS"],
    ["marks", "MARKS AND NUMBERS"],
    ["description", "KIND OF PACKAGES / DESCRIPTION OF GOODS"],
    ["gross_weight", "GROSS WEIGHT"],
    ["measurement", "MEASUREMENT"],
    ["total_words", "TOTAL NO. OF CONTAINERS OR PACKAGES (IN WORDS)"],
    ["freight", "FREIGHT AND CHARGES / REVENUE TONS / RATE / PER / PREPAID / COLLECT"],
    ["ex_rate", "EX RATE"],
    ["prepaid_at", "PREPAID AT"],
    ["payable_at", "PAYABLE AT"],
    ["place_date_issue", "PLACE AND DATE OF ISSUE"],
    ["total_prepaid_local", "TOTAL PREPAID IN LOCAL CURRENCY"],
    ["no_originals", "NO. OF ORIGINAL B(S)/L"],
    ["laden_on_board", "DATE / LADEN ON BOARD THE VESSEL"],
    ["service_term", "CY/CY (SERVICE TERM)"],
  ];
  const FIELD_LABEL = Object.fromEntries(FIELDS);
  const RIDER_ROW = ["rider", "RIDER / ATTACHED SHEET (AS PER ATTACHED RIDER / DESCRIPTIONS)"];

  // กรอบของแต่ละช่อง [x0, top, x1, bottom] หน่วย pt (A4 = 595 x 842, จุดเริ่มมุมซ้ายบน) — ตรงกับ layouts.json
  const LAYOUTS = {
    "HAL DRAFT (system print)": {
      detect_label: "Shipper", ignore_font_below: 6.0,
      fields: {
        bl_no: [355, 28, 595, 53], shipper: [12, 15, 355, 91], consignee: [12, 91, 355, 166],
        notify: [12, 166, 355, 240], pre_carriage: [12, 240, 195, 259], place_receipt: [195, 240, 355, 259],
        vessel_voy: [12, 259, 195, 278], pol: [195, 259, 355, 278], pod: [12, 278, 195, 296],
        place_delivery: [195, 278, 372, 296], final_dest: [372, 278, 595, 296],
        total_words: [12, 544, 470, 566], service_term: [470, 544, 595, 566], freight: [12, 566, 595, 692],
        ex_rate: [12, 692, 82, 717], prepaid_at: [82, 692, 243, 717], payable_at: [243, 692, 383, 717],
        place_date_issue: [383, 692, 595, 717], total_prepaid_local: [82, 717, 243, 747],
        no_originals: [243, 717, 383, 747], laden_on_board: [12, 747, 383, 842],
      },
      cargo: { top: 296, bottom: 544, columns: { left: [12, 164], no_pkgs: [164, 228], description: [228, 393],
        gross_weight: [393, 491], measurement: [491, 595] } },
    },
    "CLIENT (shipper print)": {
      detect_label: "", ignore_font_below: 6.0,
      fields: {
        bl_no: [370, 20, 595, 50], shipper: [12, 15, 360, 87], consignee: [12, 87, 360, 164],
        notify: [12, 164, 360, 240], pre_carriage: [12, 240, 214, 263], place_receipt: [214, 240, 360, 263],
        vessel_voy: [12, 263, 214, 286], pol: [214, 263, 360, 286], pod: [12, 286, 214, 315],
        place_delivery: [214, 286, 390, 315], final_dest: [390, 286, 595, 315],
        total_words: [12, 551, 470, 573], service_term: [470, 551, 595, 573], freight: [12, 573, 595, 722],
        ex_rate: [12, 722, 103, 748], prepaid_at: [103, 722, 258, 748], payable_at: [258, 722, 385, 748],
        place_date_issue: [385, 722, 595, 748], total_prepaid_local: [103, 748, 258, 773],
        no_originals: [258, 748, 385, 773], laden_on_board: [12, 773, 385, 842],
      },
      cargo: { top: 315, bottom: 551, columns: { left: [12, 183.5], no_pkgs: [183.5, 226], description: [226, 414],
        gross_weight: [414, 502], measurement: [502, 595] } },
    },
  };

  const FORM_NOISE = new Set(["BY", "AS CARRIER", "HEUNG A LINE CO., LTD."]);
  const CONTAINER_RE = /\b([A-Z]{4})\s?(\d{6})(\d)\b/;
  const CONTAINER_G = /\b([A-Z]{4})\s?(\d{6})(\d)\b/g;
  const BL_RE = /\b[A-Z]{3,6}\d{6,}\b/;
  const NUM_G = /\d[\d,]*\.?\d*/g;
  const DATE_G = /\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)[A-Z]*[.\s/-]*(\d{1,2})[.,\s/-]*(\d{4})\b/g;
  const ATTACH_RE = new RegExp(
    "\\b(AS\\s+PER|SEE|PL(?:EA)?SE?\\s+SEE|PLS\\s+SEE|DETAILS?\\s+AS\\s+PER|CONT(?:INUED|'?D)?\\s+(?:ON|IN)|REFER\\s+TO)" +
    "\\s+(?:THE\\s+)?(?:ATTACHED|ATTACHMENT|RIDER|ATTACH)" +
    "|\\bATTACHED\\s+(?:RIDER|SHEETS?|LISTS?|DESCRIPTIONS?|PAGES?|DETAILS?)\\b");
  const RIDER_NOISE_RE = new RegExp(
    "^(PAGE\\s*\\d+(\\s*(OF|/)\\s*\\d+)?|-?\\s*\\d+\\s*-?|RIDER( PAGE)?|ATTACHED SHEET|ATTACHMENT|CONTINUED|CONT'D|" +
    "\\(?CONTINUED\\)?|HEUNG A LINE CO\\., LTD\\.|" +
    "((RIDER|ATTACHED SHEET|ATTACHMENT|CONTINUATION SHEET)( PAGE)?\\W*(TO\\W*)?)?" +
    "(B/L|BILL OF LADING)\\W*NO\\.?\\W*[A-Z]{3,6}\\d{6,}|(RIDER|ATTACHED SHEET|ATTACHMENT)\\W+[A-Z]{3,6}\\d{6,})$");
  const WEIGHT_G = /([\d,]+\.?\d*)\s*KGS?\b/g;
  const NUMBER_WORDS = ("ZERO ONE TWO THREE FOUR FIVE SIX SEVEN EIGHT NINE TEN ELEVEN TWELVE THIRTEEN FOURTEEN " +
    "FIFTEEN SIXTEEN SEVENTEEN EIGHTEEN NINETEEN TWENTY").split(" ");

  const MATCH = "MATCH", MISMATCH = "NOT MATCH", ONLY_CLIENT = "MISSING IN DRAFT", ONLY_DRAFT = "MISSING IN CLIENT",
    EMPTY = "BOTH EMPTY", FORMAT = "MATCH (format differs)", VIA_RIDER = "MATCH (via rider)";
  const isProblem = s => s === MISMATCH || s === ONLY_CLIENT || s === ONLY_DRAFT;

  // ───────────── ตัวช่วยข้อความ ─────────────
  const norm = s => String(s).toUpperCase().replace(/\s+/g, " ").trim();
  const compact = s => String(s).toUpperCase().replace(/\s/g, "");
  const join = ls => ls.join(" ");
  const fmt3 = n => n.toLocaleString("en-US", { minimumFractionDigits: 3, maximumFractionDigits: 3 });
  function numbers(s) {
    return (String(s).match(NUM_G) || []).map(m => parseFloat(m.replace(/,/g, ""))).filter(v => !isNaN(v));
  }
  function countValue(s) {
    s = String(s).toUpperCase();
    const n = s.match(/\d+/);
    if (n) return parseInt(n[0], 10);
    for (let i = 0; i < NUMBER_WORDS.length; i++) if (new RegExp("\\b" + NUMBER_WORDS[i] + "\\b").test(s)) return i;
    return null;
  }
  function dates(s) {
    return [...String(s).toUpperCase().matchAll(DATE_G)].map(m => `${m[1]} ${String(+m[2]).padStart(2, "0")} ${m[3]}`);
  }
  const attachRef = lines => { const m = norm(join(lines)).match(ATTACH_RE); return m ? m[0] : ""; };

  // diff แบบ LCS ทีละคำ → กลุ่มที่ต่างกัน [{a:[...], b:[...]}]
  function diffBlocks(a, b) {
    const n = a.length, m = b.length;
    const L = Array.from({ length: n + 1 }, () => new Int32Array(m + 1));
    for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--)
      L[i][j] = a[i] === b[j] ? L[i + 1][j + 1] + 1 : Math.max(L[i + 1][j], L[i][j + 1]);
    const blocks = [];
    let i = 0, j = 0, cur = null;
    const flush = () => { if (cur) { blocks.push(cur); cur = null; } };
    while (i < n || j < m) {
      if (i < n && j < m && a[i] === b[j]) { flush(); i++; j++; }
      else if (j < m && (i === n || L[i][j + 1] >= L[i + 1][j])) { (cur = cur || { a: [], b: [] }).b.push(b[j++]); }
      else { (cur = cur || { a: [], b: [] }).a.push(a[i++]); }
    }
    flush();
    return blocks;
  }
  function tokenDiff(aLines, bLines) {
    const blocks = diffBlocks(norm(join(aLines)).split(" ").filter(Boolean), norm(join(bLines)).split(" ").filter(Boolean));
    const onlyA = blocks.filter(x => x.a.length).map(x => `"${x.a.join(" ")}"`);
    const onlyB = blocks.filter(x => x.b.length).map(x => `"${x.b.join(" ")}"`);
    const parts = [];
    if (onlyA.length) parts.push("CLIENT has: " + onlyA.join(" | "));
    if (onlyB.length) parts.push("DRAFT has: " + onlyB.join(" | "));
    return parts.join("; ");
  }
  function similarity(a, b) {   // เหมือน difflib ratio: 2*LCS/(len a + len b)
    if (!a.length && !b.length) return 1;
    let prev = new Int32Array(b.length + 1);
    for (let i = 1; i <= a.length; i++) {
      const cur = new Int32Array(b.length + 1);
      for (let j = 1; j <= b.length; j++) cur[j] = a[i - 1] === b[j - 1] ? prev[j - 1] + 1 : Math.max(prev[j], cur[j - 1]);
      prev = cur;
    }
    return 2 * prev[b.length] / (a.length + b.length);
  }

  // ───────────── อ่าน PDF (pdf.js) → คำพร้อมพิกัด แบบเดียวกับ pdfplumber ─────────────
  async function pdfWords(pdfjsLib, buf) {
    const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
    const pages = [];
    for (let p = 1; p <= pdf.numPages; p++) {
      const page = await pdf.getPage(p);
      const vp = page.getViewport({ scale: 1 });
      const words = [];
      for (const it of (await page.getTextContent()).items) {
        if (!it.str || !it.str.trim()) continue;
        const [a, b, c, d, e, f] = it.transform;
        const size = Math.hypot(c, d) || Math.hypot(a, b);
        const [x, y] = vp.convertToViewportPoint(e, f);
        const cw = (it.width || 0) / it.str.length;
        const re = /\S+/g;
        let m;
        while ((m = re.exec(it.str))) {
          const x0 = x + m.index * cw;
          words.push({ text: m[0], x0, x1: x0 + m[0].length * cw, top: y - 0.8 * size, bottom: y + 0.2 * size, size });
        }
      }
      pages.push(words);
    }
    return pages;
  }

  function groupLines(words, tol = 3.0) {
    const lines = [];
    for (const w of [...words].sort((p, q) => p.top - q.top || p.x0 - q.x0)) {
      const last = lines[lines.length - 1];
      if (last && Math.abs(last[0].top - w.top) <= tol) last.push(w); else lines.push([w]);
    }
    return lines.map(l => l.sort((p, q) => p.x0 - q.x0));
  }
  const lineText = l => l.map(w => w.text).join(" ");
  function inBox(w, box) {
    const cx = (w.x0 + w.x1) / 2, cy = (w.top + w.bottom) / 2;
    return box[0] <= cx && cx < box[2] && box[1] <= cy && cy < box[3];
  }
  function pageLines(words, minSize = 6.0) {
    return groupLines(words.filter(w => w.size >= minSize)).map(lineText)
      .filter(l => l.trim() && !RIDER_NOISE_RE.test(norm(l)));
  }

  function extractWithLayout(words, layout) {
    words = words.filter(w => w.size >= (layout.ignore_font_below || 0));
    const out = {};
    for (const [key, box] of Object.entries(layout.fields))
      out[key] = groupLines(words.filter(w => inBox(w, box))).map(lineText).filter(l => !FORM_NOISE.has(l.trim().toUpperCase()));
    // ตารางสินค้า: ตัดแต่ละบรรทัดเป็นช่วงตามช่องว่างกว้าง แล้วใส่ช่วงลงคอลัมน์ที่มันอยู่
    const cargo = layout.cargo, cols = cargo.columns, order = Object.keys(cols);
    const table = Object.fromEntries(order.map(c => [c, []]));
    const cwords = words.filter(w => { const cy = (w.top + w.bottom) / 2; return cargo.top <= cy && cy < cargo.bottom; });
    for (const line of groupLines(cwords)) {
      const segs = [];
      let cur = [line[0]];
      for (const w of line.slice(1)) {
        if (w.x0 - cur[cur.length - 1].x1 > 6.5) { segs.push(cur); cur = [w]; } else cur.push(w);
      }
      segs.push(cur);
      const row = {};
      for (const seg of segs) {
        const x0 = seg[0].x0, x1 = seg[seg.length - 1].x1;
        let col = order[order.length - 1];
        for (let i = 0; i < order.length; i++) {
          const c = order[i];
          if (cols[c][0] <= x0 + 2 && x0 + 2 < cols[c][1]) {
            col = c;
            if (c !== "description" && i + 1 < order.length) {   // ตัวเลขชิดขวาที่เริ่มเลยเส้นคอลัมน์มานิดเดียว
              const edge = cols[c][1];
              if (edge - x0 < 10 && (x1 - edge) > (edge - x0)) col = order[i + 1];
            }
            break;
          }
        }
        (row[col] = row[col] || []).push(lineText(seg));
      }
      for (const [c, parts] of Object.entries(row)) table[c].push(parts.join("  "));
    }
    out.container_seal = table.left.filter(l => CONTAINER_RE.test(l.toUpperCase()));
    out.marks = table.left.filter(l => !CONTAINER_RE.test(l.toUpperCase()));
    for (const c of ["no_pkgs", "description", "gross_weight", "measurement"]) out[c] = table[c] || [];
    return out;
  }

  function layoutScore(f, words, layout) {
    let score = 0;
    const label = layout.detect_label;
    if (label && words.some(w => w.text === label && w.size < 6)) score += 5;
    const j = k => join(f[k] || []).toUpperCase();
    if (new RegExp("^" + BL_RE.source + "$").test(j("bl_no").replace(/ /g, ""))) score += 3;
    if (CONTAINER_RE.test(j("container_seal"))) score += 2;
    if (j("total_words").includes("SAY") || j("total_words").includes("ONLY")) score += 1;
    if (dates(j("place_date_issue")).length) score += 1;
    if (/\d/.test(j("gross_weight"))) score += 1;
    if ((f.vessel_voy || []).length) score += 1;
    if ((f.shipper || []).length) score += 1;
    return score;
  }

  // pages = ผลจาก pdfWords(); riderPages = หน้าของไฟล์ RIDER แยก (ถ้ามี)
  function extractBl(name, pages, riderName, riderPages) {
    if (!pages.length || !pages[0].length)
      throw new Error(`${name}: ไม่พบตัวอักษรใน PDF (อาจเป็นไฟล์สแกนเป็นรูปภาพ) — ต้องใช้ PDF ที่เลือกข้อความได้`);
    const words = pages[0];
    let rider = [];
    for (const pg of pages.slice(1)) rider = rider.concat(pageLines(pg));
    for (const pg of riderPages || []) rider = rider.concat(pageLines(pg));
    let best = null;
    for (const [lname, lay] of Object.entries(LAYOUTS)) {
      const f = extractWithLayout(words, lay), s = layoutScore(f, words, lay);
      if (!best || s > best.score) best = { score: s, layout: lname, fields: f };
    }
    const refs = Object.entries(best.fields).filter(([, v]) => attachRef(v)).map(([k]) => FIELD_LABEL[k]);
    return { file: name, riderFile: riderName || "", layout: best.layout, score: best.score, pages: pages.length,
      fields: best.fields, rider, riderRefs: refs };
  }

  // ───────────── เทียบ ─────────────
  function compareField(key, aLines, bLines) {
    const a = norm(join(aLines)), b = norm(join(bLines));
    if (!a && !b) return [EMPTY, ""];
    if (a && !b) return [ONLY_CLIENT, `CLIENT shows "${aLines.join(" / ")}" but DRAFT is blank`];
    if (b && !a) return [ONLY_DRAFT, `DRAFT shows "${bLines.join(" / ")}" but CLIENT is blank`];
    if (a === b) return [MATCH, ""];
    if (key === "gross_weight" || key === "measurement") {
      const na = numbers(a), nb = numbers(b);
      if (na.length && na.join() === nb.join()) return [FORMAT, "Same value, written differently"];
      if (na.length && nb.length) {
        const d = na[0] - nb[0], unit = key === "gross_weight" ? "KGS" : "CBM";
        return [MISMATCH, `CLIENT ${a} vs DRAFT ${b}  (difference ${d > 0 ? "+" : ""}${fmt3(d)} ${unit})`];
      }
    }
    if (key === "no_originals") {
      const ca = countValue(a), cb = countValue(b);
      if (ca !== null && ca === cb) return [FORMAT, `Both = ${ca}  ("${a}" vs "${b}")`];
      return [MISMATCH, `CLIENT "${a}" (=${ca}) vs DRAFT "${b}" (=${cb})`];
    }
    if (compact(a) === compact(b)) return [FORMAT, "Only spacing differs"];
    return [MISMATCH, tokenDiff(aLines, bLines)];
  }

  // B/L ฝั่งหนึ่งเขียน "AS PER ATTACHED RIDER" — ทุกอย่างที่อีกฝั่งพิมพ์ในช่องนี้ต้องหาเจอในช่องนั้นหรือใน RIDER
  function checkAgainstRider(key, refSide, refLines, refRider, otherSide, otherLines) {
    const phrase = attachRef(refLines);
    if (!refRider.length)
      return [MISMATCH, `${refSide} says "${phrase}" but NO rider / attached page was found in the ${refSide} PDF`];
    const pool = compact(join(refLines.concat(refRider)));
    let missing = otherLines.filter(l => !pool.includes(compact(l)) && !attachRef([l]));
    if (key === "gross_weight" || key === "measurement") {
      const have = new Set(numbers(join(refLines.concat(refRider))));
      missing = otherLines.filter(l => !attachRef([l]) && numbers(l).some(n => !have.has(n)));
    }
    if (!missing.length)
      return [VIA_RIDER, `${refSide} says "${phrase}"; everything ${otherSide} shows here is found in the ${refSide} rider`];
    return [MISMATCH, `${refSide} says "${phrase}"; NOT found in ${refSide} rider: ` + missing.map(l => `"${l}"`).join(" | ")];
  }

  function riderContainers(lines) {
    const out = {};
    for (const l of lines) for (const m of norm(l).matchAll(CONTAINER_G)) (out[m[1] + m[2] + m[3]] = out[m[1] + m[2] + m[3]] || []).push(norm(l));
    return out;
  }
  const sumKgs = lines => [...norm(join(lines)).matchAll(WEIGHT_G)]
    .map(m => m[1].replace(/,/g, "")).filter(Boolean).map(parseFloat).filter(v => !isNaN(v)).reduce((s, v) => s + v, 0);

  // เทียบ RIDER ของ CLIENT (a) กับ DRAFT (b); page1A/page1B = ข้อความหน้า 1 ของแต่ละฝั่ง
  function compareRider(a, b, page1A = "", page1B = "") {
    if (!a.length && !b.length) return [EMPTY, "No rider / attached page on either B/L", []];
    for (const [side, other, lines, pool0, otherLines] of [["CLIENT", "DRAFT", a, page1B, b], ["DRAFT", "CLIENT", b, page1A, a]]) {
      if (lines.length && !otherLines.length) {
        const pool = compact(pool0);
        const missing = lines.filter(l => !pool.includes(compact(l))).map(norm);
        if (!missing.length)
          return [VIA_RIDER, `Only ${side} has a rider; all ${lines.length} rider line(s) are printed on ${other} page 1`, []];
        const st = side === "CLIENT" ? ONLY_CLIENT : ONLY_DRAFT;
        const diffs = missing.map(l => ({ client: side === "CLIENT" ? l : "", draft: side === "DRAFT" ? l : "", status: st,
          detail: `In ${side} rider, not found anywhere on ${other}` }));
        return [st, `${other} has NO rider; ${missing.length} of ${lines.length} ${side} rider line(s) not found on ${other} page 1: ` +
          missing.slice(0, 3).join(" | ") + (missing.length > 3 ? " ..." : ""), diffs];
      }
    }
    const ja = compact(join(a) + " " + page1A), jb = compact(join(b) + " " + page1B);
    if (norm(join(a)) === norm(join(b))) return [MATCH, `Rider text is the same (${a.length} lines)`, []];
    const diffs = [];
    if (ja !== jb) {
      const onlyA = a.filter(l => !jb.includes(compact(l))).map(norm);
      const onlyB = b.filter(l => !ja.includes(compact(l))).map(norm);
      for (const la of onlyA) {
        let best = null, score = 0.6;
        for (const lb of onlyB) { const r = similarity(la, lb); if (r > score) { best = lb; score = r; } }
        if (best) {
          onlyB.splice(onlyB.indexOf(best), 1);
          diffs.push({ client: la, draft: best, status: MISMATCH, detail: tokenDiff([la], [best]) });
        } else diffs.push({ client: la, draft: "", status: ONLY_CLIENT, detail: "Line not in DRAFT rider" });
      }
      for (const lb of onlyB) diffs.push({ client: "", draft: lb, status: ONLY_DRAFT, detail: "Line not in CLIENT rider" });
      const ca = riderContainers(a), cb = riderContainers(b);
      for (const c of [...new Set([...Object.keys(ca), ...Object.keys(cb)])].sort()) {
        if (!cb[c]) diffs.push({ client: c, draft: "", status: ONLY_CLIENT, detail: `Container ${c} is in CLIENT rider only` });
        else if (!ca[c]) diffs.push({ client: "", draft: c, status: ONLY_DRAFT, detail: `Container ${c} is in DRAFT rider only` });
      }
      const wa = sumKgs(a), wb = sumKgs(b);
      if (wa && wb && Math.abs(wa - wb) > 0.0005)
        diffs.push({ client: fmt3(wa) + " KGS", draft: fmt3(wb) + " KGS", status: MISMATCH,
          detail: `Sum of all KGS figures on the rider differs by ${wa - wb > 0 ? "+" : ""}${fmt3(wa - wb)} KGS` });
    }
    if (!diffs.length) return [FORMAT, "Same rider content, only line breaks / spacing differ", []];
    const first = diffs.slice(0, 3).map(d => d.detail || (d.client ? "CLIENT: " + d.client : "DRAFT: " + d.draft)).join("; ");
    return [MISMATCH, `${diffs.length} difference(s) in rider - see RIDER CHECK. ${first}${diffs.length > 3 ? " ..." : ""}`, diffs];
  }

  function iso6346Ok(code) {
    const m = code.match(/^([A-Z]{4})(\d{6})(\d)$/);
    if (!m) return null;
    const vals = {};
    let v = 10;
    for (const ch of "ABCDEFGHIJKLMNOPQRSTUVWXYZ") { if (v % 11 === 0) v++; vals[ch] = v++; }
    const body = m[1] + m[2];
    let total = 0;
    for (let i = 0; i < body.length; i++) total += (/[A-Z]/.test(body[i]) ? vals[body[i]] : +body[i]) * 2 ** i;
    return (total % 11) % 10 === +m[3];
  }

  const MON = { JAN: 0, FEB: 1, MAR: 2, APR: 3, MAY: 4, JUN: 5, JUL: 6, AUG: 7, SEP: 8, OCT: 9, NOV: 10, DEC: 11 };
  const toDate = s => { const [mo, d, y] = s.split(" "); return new Date(+y, MON[mo], +d); };

  function extraChecks(client, draft) {
    const notes = [];
    for (const [side, doc] of [["CLIENT", client], ["DRAFT", draft]]) {
      const f = doc.fields, txt = k => norm(join(f[k] || []));
      for (const m of txt("container_seal").matchAll(CONTAINER_G)) {
        const code = m[1] + m[2] + m[3];
        if (!iso6346Ok(code)) notes.push([side, `Container ${code} fails the ISO 6346 check digit - please verify the number`]);
      }
      const gw = numbers(txt("gross_weight")), nw = txt("description").match(/N\.?\s?W\.?\s*[:]?\s*([\d,]+\.?\d*)/);
      if (gw.length && nw) {
        const nwv = parseFloat(nw[1].replace(/,/g, ""));
        if (gw[0] < nwv) notes.push([side, `GROSS WEIGHT ${fmt3(gw[0])} KGS is LESS than N.W. ${fmt3(nwv)} KGS in the description`]);
      }
      if (doc.riderRefs.length && !doc.rider.length)
        notes.push([side, `"AS PER ATTACHED" found in ${doc.riderRefs.join(", ")} but the PDF has NO rider / attached page`]);
      for (const m of norm(join(doc.rider)).matchAll(CONTAINER_G)) {
        const code = m[1] + m[2] + m[3];
        if (!iso6346Ok(code)) notes.push([side, `Rider container ${code} fails the ISO 6346 check digit - please verify`]);
      }
      const all = txt("container_seal") + " " + norm(join(doc.rider));
      const nCntr = new Set([...all.matchAll(CONTAINER_G)].map(m => m[1] + m[2] + m[3])).size;
      const words = txt("total_words");
      let cw = null;
      for (let i = 0; i < NUMBER_WORDS.length; i++) if (new RegExp("\\b" + NUMBER_WORDS[i] + "\\b").test(words)) { cw = i; break; }
      if (cw !== null && words.includes("CONTAINER") && nCntr && cw !== nCntr)
        notes.push([side, `Total in words says ${cw} container(s) but ${nCntr} container number(s) are listed`]);
      const bl = compact(txt("bl_no")), fbl = doc.file.toUpperCase().match(BL_RE);
      if (fbl && bl && fbl[0] !== bl) notes.push([side, `File name has B/L ${fbl[0]} but the document shows ${bl}`]);
      const ob = dates(txt("laden_on_board")), iss = dates(txt("place_date_issue"));
      if (ob.length && iss.length && toDate(ob[0]) > toDate(iss[0]))
        notes.push([side, `On-board date ${ob[0]} is AFTER the issue date ${iss[0]}`]);
      if (doc.score < 5) notes.push([side, `Form layout not recognised well (layout '${doc.layout}') - results may be incomplete`]);
    }
    return notes;
  }

  // client / draft = ผลจาก extractBl()
  function checkPair(client, draft) {
    const rows = FIELDS.map(([key, label], i) => {
      const a = client.fields[key] || [], b = draft.fields[key] || [];
      let [status, detail] = compareField(key, a, b);
      const ra = attachRef(a), rb = attachRef(b);
      if (isProblem(status) && (ra || rb) && !(ra && rb)) {
        [status, detail] = ra ? checkAgainstRider(key, "CLIENT", a, client.rider, "DRAFT", b)
          : checkAgainstRider(key, "DRAFT", b, draft.rider, "CLIENT", a);
      }
      return { no: i + 1, key, field: label, client: a, draft: b, status, detail };
    });
    const page1 = doc => Object.values(doc.fields).map(join).join(" ");
    let [status, detail, riderDiffs] = compareRider(client.rider, draft.rider, page1(client), page1(draft));
    const refs = [...new Set(client.riderRefs.concat(draft.riderRefs))].sort();
    if (refs.length && status === EMPTY) {
      status = MISMATCH;
      detail = `"AS PER ATTACHED" is written in ${refs.join(", ")} but neither PDF has a rider page`;
    } else if (refs.length) detail = `Referred to in: ${refs.join(", ")}. ${detail}`;
    rows.push({ no: FIELDS.length + 1, key: RIDER_ROW[0], field: RIDER_ROW[1], client: client.rider, draft: draft.rider, status, detail });
    return { client, draft, rows, riderDiffs, notes: extraChecks(client, draft), checkedAt: new Date() };
  }

  function blNumber(res) {
    for (const side of ["draft", "client"]) { const bl = compact(join(res[side].fields.bl_no || [])); if (bl) return bl; }
    return "UNKNOWN";
  }

  // ชื่อไฟล์บอกฝั่ง: มี DRAFT (หรือ HAL ที่ไม่ใช่ส่วนของเลข HASL…) = DRAFT, มี RIDER/ATTACH = ไฟล์ RIDER แยก
  const isDraftName = n => { const u = n.toUpperCase(); return u.includes("DRAFT") || u.replace(/HASL/g, "").includes("HAL"); };
  const isRiderName = n => /RIDER|ATTACH/i.test(n);

  root.BLCheck = {
    FIELDS, RIDER_ROW, LAYOUTS, MATCH, MISMATCH, ONLY_CLIENT, ONLY_DRAFT, EMPTY, FORMAT, VIA_RIDER,
    isProblem, pdfWords, extractBl, checkPair, blNumber, isDraftName, isRiderName,
    _t: { norm, compact, tokenDiff, similarity, iso6346Ok, attachRef, compareRider, compareField },
  };
})(typeof window !== "undefined" ? window : globalThis);
