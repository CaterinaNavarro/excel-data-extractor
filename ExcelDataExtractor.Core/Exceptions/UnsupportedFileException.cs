using ExcelDataExtractor.Core.Constants;

namespace ExcelDataExtractor.Core.Exceptions;

/// <summary>
/// Exception thrown when trying to process an unsupported file.
/// </summary>
public class UnsupportedFileException : FileException
{
    public UnsupportedFileException() : base(ExceptionMessages.UnsupportedFileType)
    {
    }
}
