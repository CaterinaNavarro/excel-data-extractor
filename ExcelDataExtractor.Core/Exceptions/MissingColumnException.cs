namespace ExcelDataExtractor.Core.Exceptions;

/// <summary>
/// Exception thrown when the sheet does not contains all columns names given by
/// <c>ExcelField, ExcelSheetField or ExcelFieldAttribute </c>.
/// </summary>
public class MissingColumnException : ColumnException
{
    public MissingColumnException(string message) : base(message)
    {
    }
}
