namespace ExcelDataExtractor.Core.Exceptions;

/// <summary>
/// Exception thrown when given field <c>ExcelField, ExcelSheetField or ExcelFieldAttribute </c>
/// has <c> DataTypes </c> value not existing in the enum.
/// </summary>
public class ExcelFieldDataTypeNoExistsException : FieldException
{
    public ExcelFieldDataTypeNoExistsException(string message) : base(message)
    {
    }
}
