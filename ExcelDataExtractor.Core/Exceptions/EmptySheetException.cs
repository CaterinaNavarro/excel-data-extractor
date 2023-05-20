namespace ExcelDataExtractor.Core.Exceptions;

/// <summary>
/// Exception thrown when the sheet it is empty.
/// </summary>
public class EmptySheetException : SheetException
{
    public EmptySheetException(string message) : base(message)
    {
    }
}
