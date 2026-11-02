// Чтение настоящего файла data/A1.xlsx прямо при запуске сайта.
// Никакие задания не хранятся в JavaScript: этот модуль только разбирает XLSX.
export const EXAM_NAMES = {1:"ЦТ",2:"ЦЭ",3:"РТ",4:"ДРТ"};

const textDecoder = new TextDecoder("utf-8");
const u16 = (v, p) => v.getUint16(p, true);
const u32 = (v, p) => v.getUint32(p, true);

async function inflateRaw(bytes) {
  if (typeof DecompressionStream === "undefined") {
    throw new Error("Браузер не поддерживает распаковку XLSX. Обнови браузер до актуальной версии.");
  }
  try {
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  } catch (e) {
    throw new Error("Не удалось распаковать содержимое Excel-файла: " + e.message);
  }
}

async function unzipXlsx(buffer) {
  const bytes = new Uint8Array(buffer);
  const view = new DataView(buffer);
  let eocd = -1;
  for (let p = bytes.length - 22; p >= Math.max(0, bytes.length - 65558); p--) {
    if (u32(view, p) === 0x06054b50) { eocd = p; break; }
  }
  if (eocd < 0) throw new Error("Файл A1.xlsx не похож на корректную книгу Excel (.xlsx).");
  const count = u16(view, eocd + 10);
  let p = u32(view, eocd + 16);
  const files = new Map();
  for (let i = 0; i < count; i++) {
    if (u32(view, p) !== 0x02014b50) throw new Error("Повреждён каталог XLSX-файла.");
    const method = u16(view, p + 10);
    const compressedSize = u32(view, p + 20);
    const nameLength = u16(view, p + 28);
    const extraLength = u16(view, p + 30);
    const commentLength = u16(view, p + 32);
    const localOffset = u32(view, p + 42);
    const name = textDecoder.decode(bytes.slice(p + 46, p + 46 + nameLength));
    const localNameLength = u16(view, localOffset + 26);
    const localExtraLength = u16(view, localOffset + 28);
    const dataStart = localOffset + 30 + localNameLength + localExtraLength;
    const compressed = bytes.slice(dataStart, dataStart + compressedSize);
    let content;
    if (method === 0) content = compressed;
    else if (method === 8) content = await inflateRaw(compressed);
    else throw new Error(`Неизвестный метод сжатия XLSX (${method}) в ${name}.`);
    files.set(name, content);
    p += 46 + nameLength + extraLength + commentLength;
  }
  return files;
}

function parseXml(bytes, fileName) {
  if (!bytes) return null;
  const xml = new DOMParser().parseFromString(textDecoder.decode(bytes), "application/xml");
  if (xml.querySelector("parsererror")) throw new Error(`Не удалось прочитать XML внутри ${fileName}.`);
  return xml;
}
function colIndex(cellRef) {
  const letters = (cellRef.match(/^[A-Z]+/i) || [""])[0].toUpperCase();
  let n = 0;
  for (const ch of letters) n = n * 26 + ch.charCodeAt(0) - 64;
  return n - 1;
}
function cellValue(cell, shared) {
  const type = cell.getAttribute("t");
  if (type === "inlineStr") return [...cell.querySelectorAll("is t")].map(x => x.textContent).join("");
  const v = cell.querySelector("v");
  if (!v) return "";
  const raw = v.textContent;
  if (type === "s") return shared[Number(raw)] ?? "";
  if (type === "b") return raw === "1";
  if (type === "str") return raw;
  const n = Number(raw);
  return raw !== "" && Number.isFinite(n) ? n : raw;
}

async function readWorkbook(buffer) {
  const files = await unzipXlsx(buffer);
  const sharedXml = parseXml(files.get("xl/sharedStrings.xml"), "sharedStrings.xml");
  const shared = sharedXml ? [...sharedXml.getElementsByTagName("si")].map(si =>
    [...si.getElementsByTagName("t")].map(t => t.textContent).join("")
  ) : [];

  let sheetPath = "xl/worksheets/sheet1.xml";
  const workbook = parseXml(files.get("xl/workbook.xml"), "workbook.xml");
  const rels = parseXml(files.get("xl/_rels/workbook.xml.rels"), "workbook.xml.rels");
  const firstSheet = workbook?.getElementsByTagName("sheet")[0];
  if (firstSheet && rels) {
    const relId = firstSheet.getAttribute("r:id");
    const relation = [...rels.getElementsByTagName("Relationship")].find(x => x.getAttribute("Id") === relId);
    if (relation) {
      const target = relation.getAttribute("Target").replace(/^\//, "");
      sheetPath = target.startsWith("xl/") ? target : "xl/" + target.replace(/^\.\//, "");
    }
  }
  const sheet = parseXml(files.get(sheetPath), sheetPath);
  if (!sheet) throw new Error(`В книге не найден лист ${sheetPath}.`);
  const rows = [];
  for (const rowNode of sheet.getElementsByTagName("row")) {
    const row = [];
    for (const cell of rowNode.getElementsByTagName("c")) row[colIndex(cell.getAttribute("r"))] = cellValue(cell, shared);
    rows.push(row);
  }
  return rows;
}

function rowsToTasks(rows) {
  const headerIndex = rows.findIndex(r => r.some(v => String(v ?? "").trim() === "YEAR") && r.some(v => String(v ?? "").trim() === "WORD"));
  if (headerIndex < 0) throw new Error("В A1.xlsx не найдена строка заголовков YEAR и WORD.");
  const header = rows[headerIndex].map(v => String(v ?? "").trim());
  const ix = Object.fromEntries(["YEAR","EX","V","A(0)/O(1)","TIPE","№","BOOL","WORD","DIFF"].map(k => [k, header.indexOf(k)]));
  for (const k of ["YEAR","EX","V","TIPE","№","BOOL","WORD"]) if (ix[k] < 0) throw new Error(`В таблице отсутствует обязательный столбец «${k}».`);

  const tasks = [];
  let current = null;
  for (const row of rows.slice(headerIndex + 1)) {
    const variant = row[ix.V];
    if (variant !== undefined && variant !== null && String(variant).trim() !== "") {
      current = {
        year: Number(row[ix.YEAR]),
        exam: Number(row[ix.EX]),
        variant: String(variant).trim(),
        ao: ix["A(0)/O(1)"] >= 0 ? Number(row[ix["A(0)/O(1)"]] ?? 0) : 0,
        difficulty: ix.DIFF >= 0 ? Number(row[ix.DIFF] ?? 1) : 1,
        words: []
      };
      tasks.push(current);
    }
    if (!current) continue;
    const rawWord = row[ix.WORD];
    if (rawWord === undefined || rawWord === null || String(rawWord).trim() === "") continue;
    const correctFlag = Number(row[ix.BOOL]);
    const sourceText = String(rawWord).trim();
    const correct = correctFlag === 1 ? "А" : "О";
    current.words.push({
      position: Number(row[ix["№"]]) || current.words.length + 1,
      type: String(row[ix.TIPE] ?? "").trim().toUpperCase(),
      sourceText,
      display: sourceText.replace(/\s*\([АО]\)\s*$/u, ""),
      correct
    });
  }
  return tasks.filter(t => t.words.length).map((t, i) => ({
    ...t,
    id: `${t.year}_${t.exam}_${t.variant}_${i + 1}`
  }));
}

export async function loadDatabase() {
  const url = new URL("../data/A1.xlsx", import.meta.url);
  const response = await fetch(url, { cache: "no-store" });
  if (!response.ok) throw new Error(`Не удалось загрузить data/A1.xlsx (HTTP ${response.status}). Проверь, что файл лежит в папке data.`);
  const rows = await readWorkbook(await response.arrayBuffer());
  const tasks = rowsToTasks(rows);
  if (!tasks.length) throw new Error("Файл Excel прочитан, но в нём не найдено ни одного задания. Проверь заполнение столбцов V и WORD.");
  return tasks;
}
