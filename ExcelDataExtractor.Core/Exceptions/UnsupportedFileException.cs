using ExcelDataExtractor.Core.Constants;

namespace ExcelDataExtractor.Core.Exceptions;

/// <summary>
/// Exception thrown when trying to process an unsupported file.
/// </summary>
public class UnsupportedFileException : Exception
{
    public UnsupportedFileException() : base(ExceptionMessages.UnsupportedFileType)
    {
    }
}
