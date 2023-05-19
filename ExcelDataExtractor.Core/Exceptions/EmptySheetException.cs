namespace ExcelDataExtractor.Core.Exceptions;

public class EmptySheetException : Exception
{
    public EmptySheetException(string message) : base(message)
    {
    }
}
