export default (workbook) => {
  const sheet = workbook.sheet(0);
  sheet.cell("A1").value("TEST");
};
