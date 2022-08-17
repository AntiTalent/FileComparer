using ExcelComparer;

// example to use:

var folder = Path.Combine("C:", "work");

var excelFile1 = Path.Combine(folder, "file1.xlsx");
var excelFile2 = Path.Combine(folder, "file2.xlsx");
var excelComparer = new ComparerEPPlus();
excelComparer.Compare(excelFile1, excelFile2);

var xmlFile1 = Path.Combine(folder, "file1.xlsx");
var xmlFile2 = Path.Combine(folder, "file2.xlsx");
var xmlComparer = new ComparerXml();
xmlComparer.Compare(xmlFile1, xmlFile2);

