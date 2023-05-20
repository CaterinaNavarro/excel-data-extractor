namespace ExcelDataExtractor.Core.Exceptions;

/// <summary>
/// Exception thrown when the field type value is different from the one given by <c> DataTypes </c> value.
/// </summary>
public class FieldValueTypeDifferentFieldDataTypeException : FieldException
{
    public FieldValueTypeDifferentFieldDataTypeException(string message) : base(message)
    {
    }
}
