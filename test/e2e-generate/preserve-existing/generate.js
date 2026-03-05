export default (workbook) => {
  workbook.sheets().forEach((sheet) => {
    sheet.cell("A1").value("FOO");
  });
};
