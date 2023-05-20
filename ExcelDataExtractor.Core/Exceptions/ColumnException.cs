namespace ExcelDataExtractor.Core.Exceptions;

/// <summary>
/// Class inherited by exceptions related exclusively to columns.
/// </summary>
public abstract class ColumnException : Exception
{
    public ColumnException(string message) : base(message)
    {
    }
}
