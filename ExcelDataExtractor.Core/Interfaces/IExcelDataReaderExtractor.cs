using ExcelDataExtractor.Core.Models;

namespace ExcelDataExtractor.Core.Interfaces
{
    /// <summary>
    /// Interface for processing, validating and extracting data from files in Excel format. Supports field validations and conversion to a specific type.
    /// </summary>
    public interface IExcelDataReaderExtractor
    {
        /// <summary>
        /// Extract all data of all sheets.
        /// </summary>
        /// <param name="byteArrayContent"> Byte array content. </param>
        /// <param name="excludeSheetsWithNoneOrOneRows"> If <c> True </c> exclude sheets with none or one rows, 
        /// if <c>False</c> the result could contain any <c>IEnumerable</c> with no <c>Dictionary</c> items. </param>
        /// <returns> <para>An <c> IEnumerable </c> where each item represents a sheet. 
        /// Each sheet contains a sequence of <c>Dictionary</c>, a single <c>Dictionary</c> represents only one row of the sheet.</para>
        /// The key of the <c>Dictionary</c> is the column name, and the value is the stored on the current field.</returns>
        IEnumerable<IEnumerable<Dictionary<string, object?>>> ProcessExtractData(byte[] byteArrayContent, bool excludeSheetsWithNoneOrOneRows);

        /// <summary>
        /// Extract specific data, performs fields validations.
        /// </summary>
        /// <param name="byteArrayContent"> Byte array content. </param>
        /// <param name="fields"> Fields that the sheets must contain. </param>
        /// <param name="ignoreUnindicatedFields"> <para>If <c>true</c> does not make any validations on the fields that exists in the file but were not indicated as fields,
        /// as consequence it does not extract them neither.</para>
        /// If <c>false</c> validate the sheets contains the columns indicated only.</param>
        /// <param name="excludeSheetsWithNoneOrOneRows"> If <c> True </c> exclude sheets with none or one rows, 
        /// if <c>False</c> the result could contain any <c>IEnumerable</c> with no <c>Dictionary</c> items. </param>
        /// <returns> <para>An <c> IEnumerable </c> where each item represents a sheet. 
        /// Each sheet contains a sequence of <c>Dictionary</c>, a single <c>Dictionary</c> represents only one row of the sheet.</para>
        /// The key of the <c>Dictionary</c> is the column name, and the value is the stored on the current field.</returns>
        IEnumerable<IEnumerable<Dictionary<string, object?>>> ProcessExtractData(byte[] byteArrayContent, IEnumerable<ExcelSheetField> fields, bool ignoreUnindicatedFields, bool excludeSheetsWithNoneOrOneRows);

        /// <summary>
        /// Extract the data of a specific sheet, performing fields validations.
        /// </summary>
        /// <typeparam name="T"> Output class whose properties contains <c>JsonPropertyAttribute</c> (or another, if necessary) for matching the columns names. </typeparam>
        /// <param name="byteArrayContent"> Byte array content. </param>
        /// <param name="fields"> Fields that the sheet must contain. </param>
        /// <param name="ignoreUnindicatedFields"> <para>If <c>true</c> does not make any validations on the fields that exists in the sheet but were not indicated as fields,
        /// as consequence it does not extract them neither.</para>
        /// If <c>false</c> validate the sheet contains the columns indicated only.</param>
        /// <param name="sheetIndex"> Sheet index to extract, as default is the first. </param>
        /// <returns> An <c> IEnumerable </c> containing the rows converted into the output type. </returns>
        IEnumerable<T> ProcessExtractDataSheet<T>(byte[] byteArrayContent, IEnumerable<ExcelField> fields, bool ignoreUnindicatedFields, int sheetIndex = 0);

        /// <summary>
        /// Extract the data of a specific sheet, performing fields validations.
        /// </summary>
        /// <typeparam name="T"> Output class whose properties contains the <c>ExcelFieldAttribute</c> for matching the columns names and provide specific information of the fields. </typeparam>
        /// <param name="byteArrayContent"> Byte array content. </param>
        /// <param name="ignoreUnindicatedFields"> <para>If <c>true</c> does not make any validations on the fields that exists in the sheet but were not indicated as fields,
        /// as consequence it does not extract them neither.</para>
        /// If <c>false</c> validate the sheet contains the columns indicated only.</param>
        /// <param name="sheetIndex"> Sheet index to extract, as default it is the first. </param>
        /// <returns> An <c> IEnumerable </c> containing the rows converted into the output type. </returns>
        IEnumerable<T> ProcessExtractDataSheet<T>(byte[] byteArrayContent, bool ignoreUnindicatedFields, int sheetIndex = 0);
    }
}
