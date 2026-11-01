// Загрузка Excel-базы из папки data. SheetJS подключается из CDN.
export const EXAM_NAMES = {1:"ЦТ",2:"ЦЭ",3:"РТ",4:"ДРТ"};
const EXCEL_URL = "./data/A1.xlsx";

function cell(row, index) { return row?.[index] ?? ""; }
function normalizeType(v) {
  const s = String(v ?? "").trim().toUpperCase();
  return ["P","C","N"].includes(s) ? s : "N";
}
function parseWord(raw, boolValue, position, type) {
  const text = String(raw ?? "").trim();
  if (!text) return null;
  const answerMatch = text.match(/\(([АО])\)/i);
  const correct = answerMatch ? answerMatch[1].toUpperCase() : (Number(boolValue) === 1 ? "А" : "О");
  return {
    position: Number(position) || 0,
    type: normalizeType(type),
    sourceText: text,
    display: text.replace(/\(([АО])\)/ig, ""),
    correct
  };
}
export function parseWorkbook(workbook, XLSX) {
  const tasks = [];
  workbook.SheetNames.forEach(sheetName => {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, {header:1, defval:"", raw:true});
    let carryYear = "", carryExam = "", carryVariant = "", carryAO = "";
    let group = null;
    const flush = () => {
      if (group && group.words.length) {
        if (group.words.length === 5) {
          tasks.push({...group, id:`${group.year}_${group.exam}_${group.variant}_${tasks.length+1}`});
        } else {
          console.warn("Пропущена неполная группа задания:", group);
        }
      }
      group = null;
    };
    for (let i=0;i<rows.length;i++) {
      const row = rows[i];
      const yearCell=cell(row,1), examCell=cell(row,2), variantCell=cell(row,3), aoCell=cell(row,4);
      if (yearCell !== "") carryYear=yearCell;
      if (examCell !== "") carryExam=examCell;
      if (variantCell !== "") {
        flush();
        carryVariant=variantCell;
        carryAO=aoCell;
        group={year:Number(carryYear)||2021,exam:Number(carryExam)||1,variant:String(carryVariant),
          ao:Number(carryAO)||0,words:[],sourceSheet:sheetName,sourceRow:i+1};
      } else if (!group && cell(row,8)) {
        // Игнорируем строки заголовка/легенды, пока не встретится номер варианта.
        continue;
      }
      if (!group) continue;
      const rawWord=cell(row,8);
      if (!rawWord) continue;
      const word=parseWord(rawWord,cell(row,7),cell(row,6),cell(row,5));
      if (word) group.words.push(word);
      if (group.words.length===5) flush();
    }
    flush();
  });
  return tasks;
}
export async function loadDatabase() {
  if (!window.XLSX) throw new Error("Библиотека SheetJS не загрузилась. Проверь подключение к интернету.");
  const response = await fetch(EXCEL_URL, {cache:"no-store"});
  if (!response.ok) throw new Error(`Не удалось открыть ${EXCEL_URL} (HTTP ${response.status}).`);
  const buffer = await response.arrayBuffer();
  const workbook = window.XLSX.read(buffer, {type:"array"});
  return parseWorkbook(workbook, window.XLSX);
}
