namespace ExcelDataExtractor.Core.Exceptions
{
    public abstract class FieldException : Exception
    {
        public FieldException(string message) : base(message)
        {
        }
    }
}
