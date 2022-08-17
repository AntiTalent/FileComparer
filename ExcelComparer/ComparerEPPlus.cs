using OfficeOpenXml;

namespace ExcelComparer;

/// <summary>
///   Compare Excel cell values across all sheets using EPPlus
/// </summary>
public class ComparerEPPlus : IFileComparer
{
    public ComparerEPPlus()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }
    
    public void Compare(string file1, string file2)
    {
        Console.WriteLine(
            $"\nComparing {file1} => {file2}");

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var pckgOriginal = new ExcelPackage(new FileInfo(file1)); // IDisposable
        var pckgRepaired = new ExcelPackage(new FileInfo(file2)); // IDisposable
        var wbOriginal = pckgOriginal.Workbook; // IDisposable
        var wbRepaired = pckgRepaired.Workbook; // IDisposable

        foreach (var wsOriginal in wbOriginal.Worksheets) // IDisposable
        {
            if (wsOriginal.Dimension == null) // empty sheet
            {
                wsOriginal.Dispose();
                continue;
            }

            Console.WriteLine(wsOriginal.Name);

            var wsRepaired = wbRepaired.Worksheets[wsOriginal.Name]; // IDisposable

            const int startRow = 1;
            const int startCol = 1;
            var endRow = wsOriginal.Dimension.Rows;
            var endCol = wsOriginal.Dimension.Columns;

            for (var row = startRow; row <= endRow; ++row)
            {
                for (var col = startCol; col <= endCol; ++col)
                {
                    if (!CompareExcelValues(wsOriginal.Cells[row, col].Value, wsRepaired.Cells[row, col].Value))
                        Console.WriteLine(
                            $"\t{ExcelCellBase.GetAddress(row, col)} (value): {wsOriginal.Cells[row, col].Value} <=> {wsRepaired.Cells[row, col].Value}");
                    if (wsOriginal.Cells[row, col].FormulaR1C1 != wsRepaired.Cells[row, col].FormulaR1C1)
                        Console.WriteLine(
                            $"\t{ExcelCellBase.GetAddress(row, col)} (frmla): {wsOriginal.Cells[row, col].Formula} <=> {wsRepaired.Cells[row, col].Formula}");
                }
            }

            wsRepaired.Dispose();
            wsOriginal.Dispose();
        }

        wbRepaired.Dispose();
        wbOriginal.Dispose();

        pckgRepaired.Dispose();
        pckgOriginal.Dispose();
    }

    private static bool CompareExcelValues(object? value1, object? value2)
    {
        var s1 = value1 == null ? string.Empty : value1.ToString()!.Trim();
        var s2 = value2 == null ? string.Empty : value2.ToString()!.Trim();
        return s1 == s2;
    }
}