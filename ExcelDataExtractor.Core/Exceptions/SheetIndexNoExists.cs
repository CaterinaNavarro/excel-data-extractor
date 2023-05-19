using ExcelDataExtractor.Core.Constants;

namespace ExcelDataExtractor.Core.Exceptions
{
    public class SheetIndexNoExists : Exception
    {
        public SheetIndexNoExists() : base (ExceptionMessages.SheetIndexNoExists)
        {
        }
    }
}
