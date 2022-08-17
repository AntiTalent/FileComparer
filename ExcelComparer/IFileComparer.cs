namespace ExcelComparer;

/// <summary>
///   Compare 2 files in some way
/// </summary>
public interface IFileComparer
{
    void Compare(string file1, string file2);
}