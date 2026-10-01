import ExcelJS from "exceljs";

export async function importFixture() {
  const book = new ExcelJS.Workbook();
  const inputs = book.addWorksheet("Inputs");
  inputs.getCell("A1").value = 3;
  inputs.getCell("A2").value = 7;
  inputs.getCell("A3").value = 1234.5; inputs.getCell("A3").numFmt = '$#,##0.00';
  inputs.getCell("B1").value = new Date("2024-03-28T00:00:00Z"); inputs.getCell("B1").numFmt = 'yyyy-mm-dd';
  inputs.getCell("C1").value = { text: "Link", hyperlink: "https://example.com" };
  inputs.mergeCells("D1:E1"); inputs.getCell("D1").value = "Merged title";
  inputs.getCell("D1").font = {name:"Arial",size:16,bold:true,color:{argb:"FFFF0000"}};
  inputs.getCell("D1").fill = {type:"pattern",pattern:"solid",fgColor:{argb:"FFDDFFDD"}};
  inputs.getCell("D1").alignment = {horizontal:"center",wrapText:true};
  inputs.getColumn(1).width = 20; inputs.getRow(1).height = 30;
  inputs.getCell('A5').value='Hidden row'; inputs.getRow(5).hidden = true; inputs.getColumn(6).hidden = true;
  inputs.views = [{state:"frozen",xSplit:1,ySplit:1}];
  const summary = book.addWorksheet("Summary");
  summary.getCell("A1").value = {formula:"SUM(Inputs!A1:A2)",result:10};
  summary.getCell("A2").value = {formula:"'Inputs'!$A$1*2",result:6};
  summary.getCell("A3").value = {formula:"IF(Inputs!A1>2,TRUE,FALSE)",result:true};
  summary.getCell("A4").value = {formula:"IF(Inputs!A1>10,TRUE,FALSE)",result:false};
  summary.getCell("A5").value = {formula:"Inputs!A1-3",result:0};
  summary.getCell("B1").value = {formula:"D1+1"}; summary.getCell("D1").value = 4;
  summary.fillFormula("B2:B3", "Inputs!A1*2", [6,14]);
  summary.getCell("C2").value = {richText:[{text:"Bold",font:{bold:true}},{text:" plain"}]};
  const hidden = book.addWorksheet("Hidden", {state:"hidden"}); hidden.getCell("A1").value = 2;
  summary.getCell("A6").value = {formula:"Hidden!A1+1",result:3};
  const quoted = book.addWorksheet("Year's Data"); quoted.getCell("A1").value=5;
  summary.getCell("A7").value={formula:"'Year''s Data'!A1+Inputs!A1",result:8};
  return book.xlsx.writeBuffer();
}
