using NPOI.XSSF.UserModel;

namespace ExcelComparer;

/// <summary>
///   Save an Excel file with NPOI
/// </summary>
public class SaverNPOI : IExcelSaver
{
    public void Save(string originalFileName, string newFileName)
    {
        var fiOriginal = new FileInfo(originalFileName);
        var fiNew = new FileInfo(newFileName);
        
        using (var originalFile = File.Open(originalFileName, FileMode.Open))
        {
            var wb = new XSSFWorkbook(originalFile); // opening with path would also overwrite the original file...
            using (var fs = File.Create(newFileName))
            {
                wb.Write(fs);
            }
            wb.Close();
        }
        
        Console.WriteLine($"Original file size: {fiOriginal.Length} | New file size: {fiNew.Length} | Diff: {fiOriginal.Length - fiNew.Length}");
    }
}