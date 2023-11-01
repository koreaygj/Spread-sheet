const spreadsheetContainer = document.querySelector("#spreadsheet-container");
const exportBtn = document.querySelector("#export-btn");
const ROWS = 10;
const COLS = 10;
const spreadsheet = [];
const alphabets = [
  "A",
  "B",
  "C",
  "D",
  "E",
  "F",
  "G",
  "H",
  "I",
  "J",
  "K",
  "L",
  "M",
  "N",
  "O",
  "P",
  "Q",
  "R",
  "S",
  "T",
  "U",
  "V",
  "W",
  "X",
  "Y",
  "Z",
];

class Cell {
  constructor(
    isHeader,
    disable,
    data,
    row,
    column,
    rowName,
    colName,
    active = false
  ) {
    this.isHeader = isHeader;
    this.disable = disable;
    this.data = data;
    this.row = row;
    this.column = column;
    this.rowName = rowName;
    this.colName = colName;
    this.active = active;
  }
}

initSpreadsheet();

function initSpreadsheet() {
  for (let i = 0; i < COLS; i++) {
    let spreadsheetRow = [];
    for (let j = 0; j < ROWS; j++) {
      let cellData = "";
      let isHeader = false;
      let disable = false;
      if (j === 0) {
        cellData = i;
        isHeader = true;
        disable = true;
      }
      if (i === 0) {
        cellData = alphabets[j - 1];
        isHeader = true;
        disable = true;
      }
      if (!cellData) {
        cellData = "";
      }
      const cell = new Cell(
        isHeader,
        disable,
        cellData,
        i,
        j,
        i,
        alphabets[j - 1],
        false
      );
      spreadsheetRow.push(cell);
    }
    spreadsheet.push(spreadsheetRow);
  }
  printSheet();
}

function getElementRowCol(row, col) {
  return document.querySelector("#cell-" + row + col);
}

function createCellElement(cell) {
  const cellElement = document.createElement("input");
  cellElement.className = "cell";
  cellElement.id = `cell-${cell.row}${cell.column}`;
  cellElement.value = cell.data;
  cellElement.disabled = cell.disable;
  if (cell.isHeader) cellElement.classList.add("header");

  cellElement.onclick = () => handleCellClick(cell);
  cellElement.onchange = (e) => handleOnChange(e.target.value, cell);
  return cellElement;
}

function printSheet() {
  for (let i = 0; i < COLS; i++) {
    const rowContainer = document.createElement("div");
    rowContainer.className = "cell-row";
    for (let j = 0; j < ROWS; j++) {
      const cell = spreadsheet[i][j];
      rowContainer.appendChild(createCellElement(cell));
    }
    spreadsheetContainer.appendChild(rowContainer);
  }
}

function clearHeaderActive() {
  const headers = document.querySelectorAll(".header");
  headers.forEach((header) => header.classList.remove("active"));
}

function handleCellClick(cell) {
  clearHeaderActive();
  const columnHeader = spreadsheet[0][cell.column];
  const rowHeader = spreadsheet[cell.row][0];
  const columnHeaderElement = getElementRowCol(
    columnHeader.row,
    columnHeader.column
  );
  const rowHeaderElement = getElementRowCol(rowHeader.row, rowHeader.column);
  columnHeaderElement.classList.add("active");
  rowHeaderElement.classList.add("active");
  const cellInfo = document.querySelector("#cell-status");
  console.log(cellInfo);
  cellInfo.innerHTML = `${cell.colName} ${cell.rowName}`;
}

function handleOnChange(data, cell) {
  cell.data = data;
}

exportBtn.onclick = function (e) {
  let csv = "";
  for (let i = 1; i < spreadsheet.length; i++) {
    csv +=
      spreadsheet[i]
        .filter((item) => !item.isHeader)
        .map((item) => item.data)
        .join(",") + "\r\n";
  }

  const csvObj = new Blob([csv]);
  const csvUrl = URL.createObjectURL(csvObj);
  console.log("csv", csvUrl);

  const a = document.createElement("a");
  a.href = csvUrl;
  a.download = "spreadSheet file name.csv";
  a.click();
};
