namespace ExcelDataExtractor.Core.Exceptions
{
    public class RepeatedColumnException : ColumnException
    {
        public RepeatedColumnException(string message) : base(message)
        {
        }
    }
}
