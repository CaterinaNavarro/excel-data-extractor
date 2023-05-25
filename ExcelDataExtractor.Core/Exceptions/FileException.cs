namespace ExcelDataExtractor.Core.Exceptions;

/// <summary>
/// Abstract class inherited by exceptions related exclusively to the file.
/// </summary>
public abstract class FileException : Exception
{
    public FileException(string message) : base(message)
    {
    }
}
