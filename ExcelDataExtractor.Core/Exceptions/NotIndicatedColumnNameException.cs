namespace ExcelDataExtractor.Core.Exceptions;

/// <summary>
/// Exception thrown when <c> ignoreUnindicatedFields </c> is <c>false</c>, the process validate all the columns names of the sheet 
/// and the sheet has more columns names than the given by <c> ExcelField, ExcelSheetField or ExcelFieldAttribute </c>.
/// </summary>
public class NotIndicatedColumnNameException : ColumnException
{
    public NotIndicatedColumnNameException(string message) : base(message)
    {
    }
}
