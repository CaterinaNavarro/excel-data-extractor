using ExcelDataExtractor.Core.Constants;

namespace ExcelDataExtractor.Core.Exceptions;

/// <summary>
/// Exception thrown when the given field <c>ExcelField, ExcelSheetField or ExcelFieldAttribute </c> 
/// has no column name.
/// </summary>
public class ExcelFieldColumnNameNullEmptyException : FieldException
{
    public ExcelFieldColumnNameNullEmptyException() : base(ExceptionMessages.ExcelFieldColumnNameNullEmpty)
    {
    }
}
