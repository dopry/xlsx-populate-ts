export default (workbook) => {
  workbook
    .cloneSheet(workbook.sheet(0), "Sheet Cloned")
    .cell("A3")
    .value("baz");
};
