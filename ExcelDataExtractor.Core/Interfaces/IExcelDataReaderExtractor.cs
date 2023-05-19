using ExcelDataExtractor.Core.Models;

namespace ExcelDataExtractor.Core.Interfaces
{
    public interface IExcelDataReaderExtractor
    {
        List<List<Dictionary<string, object?>>> ProcessExtractData(byte[] byteArrayContent, IEnumerable<ExcelSheetField>? fields = null, bool? ignoreUnindicatedFields = null);
        List<T> ProcessExtractDataSheet<T>(byte[] byteArrayContent, int sheetIndex = 0);
        List<T> ProcessExtractDataSheet<T>(byte[] byteArrayContent, IEnumerable<ExcelField> fields, bool ignoreUnindicatedFields, int sheetIndex = 0);
        List<T> ProcessExtractDataSheet<T>(byte[] byteArrayContent, bool ignoreUnindicatedFields, int sheetIndex = 0);
    }
}
