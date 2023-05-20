namespace ExcelDataExtractor.Core.Exceptions;

/// <summary>
/// Abstract class inherited by exceptions related exclusively to columns.
/// </summary>
public abstract class ColumnException : Exception
{
    public ColumnException(string message) : base(message)
    {
    }
}
