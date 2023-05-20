namespace ExcelDataExtractor.Core.Exceptions;

/// <summary>
/// Exception thrown when a sheet as repeated columns names.
/// </summary>
public class RepeatedColumnException : ColumnException
{
    public RepeatedColumnException(string message) : base(message)
    {
    }
}
