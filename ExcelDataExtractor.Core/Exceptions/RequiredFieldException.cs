namespace ExcelDataExtractor.Core.Exceptions
{
    public class RequiredFieldException : FieldException
    {
        public RequiredFieldException(string message) : base(message)
        {
        }
    }
}
