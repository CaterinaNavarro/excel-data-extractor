namespace ExcelDataExtractor.Core.Exceptions
{
    public class NoIndicatedFieldException : ColumnException
    {
        public NoIndicatedFieldException(string message) : base(message)
        {
        }
    }
}
