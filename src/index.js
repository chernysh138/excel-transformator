import fs from "fs";
import xlsx from "xlsx";

const INPUT_FILE_PATH = "./files/horoshop_torgsoft.xlsx";
const OUTPUT_FILE_PATH = "./files/output.xlsx";
const OUTPUT_FILE_SHEET_NAME = "Sheet1";

// Основная функция, содержащая всю функциональность.
async function main() {
  // В этой части мы читаем файл и получаем данные.
  const buffer = readFile();
  const data = await parseExcelData(buffer); // Получить массив объектов

  // Просматриваем все строки файла Excel
  // Рядок назвал "row"
  for (const row of data) {
    if (row["Наличие"] === "+") {
      row["Наличие"] = "В наявності";
    }
    if (String(row["Название_позиции"]).startsWith("Сліпони")) {
      row["Раздел"] = "Взуття/Сліпони";
    }
    if (String(row["Название_позиции"]).startsWith("Черевики")) {
      row["Раздел"] = "Взуття/Черевики";
    }
    // row["Раздел"]=
    ChangeNameColumn(row, "Код_товара", "Артикул для отображения на сайте");
    ChangeNameColumn(row, "Идентификатор_товара", "Артикул");
    ChangeNameColumn(
      row,
      "Значение_Характеристики",
      "Название модификации (UA)"
    );
    ChangeNameColumn(row, "Цена", "Старая цена");
    ChangeNameColumn(row, "Скидка", "Скидка %");
    row["Скидка %"] = String(row["Скидка %"]).split(",")[0];
    ChangeNameColumn(row, "Значение_Характеристики_6", "Цена");
    ChangeNameColumn(row, "Описание", "Описание товара (RU)");
    ChangeNameColumn(row, "Значение_Характеристики_1", "Цвет");
    ChangeNameColumn(row, "Производитель", "Бренд");

    DeletColumn(row, "Тип_товара");
    DeletColumn(row, "Идентификатор_группы");
    DeletColumn(row, "ID_группы_разновидностей");
    DeletColumn(row, "Единица_измерения");
    DeletColumn(row, "Название_Характеристики");
    row["Розмір"] = String(row["Значение_Характеристики_5"]).split("(")[0];
    console.log(data);
  }

  // Сохранить данные в файле «output.xlsx».
  writeExcelFile(data);
}
function ChangeNameColumn(hash_name, old_name, new_name) {
  hash_name[new_name] = hash_name[old_name];
  delete hash_name[old_name];
  return hash_name[new_name];
}
function DeletColumn(hash_name, old_name) {
  delete hash_name[old_name];
}
// Вызов основной функции
main();

/// --- Утилиты ---

// Эта функция считывает двоичные данные файла и возвращает буфер.
function readFile() {
  const buffer = fs.readFileSync(INPUT_FILE_PATH);

  return buffer;
}

// Получение данных Excel с помощью библиотеки «xlsx».
async function parseExcelData(buffer) {
  const workbook = xlsx.read(buffer, { type: "buffer" }); // Получить данные рабочей тетради Excel
  const sheetName = workbook.SheetNames[0]; // Получить название первого листа в рабочей книге
  const sheet = workbook.Sheets[sheetName]; // Получить первый лист в рабочей тетради
  const jsonData = xlsx.utils.sheet_to_json(sheet); // преобразовать данные в json

  return jsonData;
}

// Записываем наш массив объектов в файл с помощью библиотеки «xlsx».
function writeExcelFile(data) {
  const worksheet = xlsx.utils.json_to_sheet(data);
  const workbook = xlsx.utils.book_new();

  xlsx.utils.book_append_sheet(workbook, worksheet, OUTPUT_FILE_SHEET_NAME);
  xlsx.writeFile(workbook, OUTPUT_FILE_PATH);
}
