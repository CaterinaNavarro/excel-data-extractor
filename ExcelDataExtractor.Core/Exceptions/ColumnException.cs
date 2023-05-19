namespace ExcelDataExtractor.Core.Exceptions;

public abstract class ColumnException : Exception
{
    public ColumnException(string message) : base(message)
    {
    }
}
