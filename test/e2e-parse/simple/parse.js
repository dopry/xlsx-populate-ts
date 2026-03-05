export default (workbook) => {
  return workbook.sheet(0).usedRange().value();
};
