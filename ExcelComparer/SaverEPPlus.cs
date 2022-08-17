using OfficeOpenXml;

namespace ExcelComparer;

/// <summary>
///   Save an Excel file with EPPlus
/// </summary>
public class SaverEPPlus : IExcelSaver
{
    public SaverEPPlus()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    public void Save(string originalFileName, string newFileName)
    {
        var fiOriginal = new FileInfo(originalFileName);
        var fiNew = new FileInfo(newFileName);
        if (fiNew.Exists) fiNew.Delete();

        using (var pckgOriginal = new ExcelPackage(fiOriginal))
        {
            pckgOriginal.SaveAs(newFileName);
        }
        
        Console.WriteLine($"Original file size: {fiOriginal.Length} | New file size: {fiNew.Length} | Diff: {fiOriginal.Length - fiNew.Length}");
    }
}