namespace ExcelComparer;

/// <summary>
///   Save an excel file using one of the supported/implemented Excel library
/// </summary>
public interface IExcelSaver
{
    void Save(string originalFileName, string newFileName);
}