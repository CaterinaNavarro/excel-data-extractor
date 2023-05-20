namespace ExcelDataExtractor.Core.Exceptions;

/// <summary>
/// Abstract class inherited by exceptions related exclusively to sheets.
/// </summary>
public abstract class SheetException : Exception
{
    public SheetException(string message) : base(message)
    {
    }
}
