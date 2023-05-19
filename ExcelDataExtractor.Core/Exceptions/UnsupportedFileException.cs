using ExcelDataExtractor.Core.Constants;

namespace ExcelDataExtractor.Core.Exceptions;

public class UnsupportedFileException : Exception
{
    public UnsupportedFileException() : base(ExceptionMessages.UnsupportedFileType)
    {
    }
}
