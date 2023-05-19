namespace ExcelDataExtractor.Core.Exceptions
{
    public class MissingColumnException : ColumnException
    {
        public MissingColumnException(string message) : base(message)
        {
        }
    }
}
