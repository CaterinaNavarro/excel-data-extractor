using ExcelDataExtractor.Core.Models;

namespace ExcelDataExtractor.Core.Interfaces
{
    public interface IExcelDataReaderExtractor
    {
        /// <summary>
        /// Extract all data of all sheets.
        /// </summary>
        /// <param name="byteArrayContent"> Byte array content. </param>
        /// <returns> <para>List of List of Dictionary, each list item of the main list represents a sheet. 
        /// Each sheet contains a list of dictionary, a dictionary represents only one row of the sheet.</para>
        /// The key of the dictionary is the column name, and the value is the stored on the current field.</returns>
        List<List<Dictionary<string, object?>>> ProcessExtractData(byte[] byteArrayContent);

        /// <summary>
        /// Extract specific data, performs fields validations.
        /// </summary>
        /// <param name="byteArrayContent"> Byte array content. </param>
        /// <param name="fields"> Fields that the sheets must contain. </param>
        /// <param name="ignoreUnindicatedFields"> <para>If <c>true</c> does not make any validations on the fields that exists in the file but were not indicated as fields,
        /// as consequence it does not extract them neither.</para>
        /// If <c>false</c> validate the sheets contains the columns indicated only.</param>
        /// <returns> <para>List of List of Dictionary, each list item of the main list represents a sheet. 
        /// Each sheet contains a list of dictionary, a dictionary represents only one row of the sheet.</para>
        /// The key of the dictionary is the column name, and the value is the stored on the current field.</returns>
        List<List<Dictionary<string, object?>>> ProcessExtractData(byte[] byteArrayContent, IEnumerable<ExcelSheetField> fields, bool ignoreUnindicatedFields);

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
        /// <returns> The rows converted into the output class list. </returns>
        List<T> ProcessExtractDataSheet<T>(byte[] byteArrayContent, IEnumerable<ExcelField> fields, bool ignoreUnindicatedFields, int sheetIndex = 0);

        /// <summary>
        /// Extract the data of a specific sheet, performing fields validations.
        /// </summary>
        /// <typeparam name="T"> Output class whose properties contains the <c>ExcelFieldAttribute</c> for matching the columns names and provide specific information of the fields. </typeparam>
        /// <param name="byteArrayContent"> Byte array content. </param>
        /// <param name="ignoreUnindicatedFields"> <para>If <c>true</c> does not make any validations on the fields that exists in the sheet but were not indicated as fields,
        /// as consequence it does not extract them neither.</para>+6252
        /// If <c>false</c> validate the sheet contains the columns indicated only.</param>
        /// <param name="sheetIndex"> Sheet index to extract, as default it is the first. </param>
        /// <returns> The rows converted into the output class list. </returns>
        List<T> ProcessExtractDataSheet<T>(byte[] byteArrayContent, bool ignoreUnindicatedFields, int sheetIndex = 0);
    }
}
