namespace ExcelDataExtractor.Core.Exceptions;

/// <summary>
/// Exception thrown when a field is required and has no value.
/// </summary>
public class RequiredFieldException : FieldException
{
    public RequiredFieldException(string message) : base(message)
    {
    }
}
