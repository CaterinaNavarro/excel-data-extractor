namespace ExcelDataExtractor.Core.Exceptions;

/// <summary>
/// Exception thrown when a field has value but its column name is missing.
/// </summary>
public class FieldHasValueNoColumnNameException : FieldException
{
    public FieldHasValueNoColumnNameException(string message) : base(message)
    {
    }
}
