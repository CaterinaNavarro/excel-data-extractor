using Aspose.Cells;
using ExcelDataExtractor.Core.Attributes;
using ExcelDataExtractor.Core.Enums;
using ExcelDataExtractor.Core.Exceptions;
using ExcelDataExtractor.Core.Extensions;
using ExcelDataExtractor.Core.Helpers;
using ExcelDataExtractor.Core.Interfaces;
using ExcelDataExtractor.Core.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.ComponentModel;
using System.Reflection;

namespace ExcelDataExtractor.Core;

/// <summary>
/// <c>IExcelDataReaderExtractor</c> implementation, for processing, validating and extracting data from files in Excel format. Supports field validations and conversion to a specific type.
/// </summary>
public class ExcelDataReaderExtractor : IExcelDataReaderExtractor
{
    private readonly TypeConverterHelper _typeConverter = new();

    /// <summary>
    /// Extract all data of all sheets.
    /// </summary>
    /// <param name="byteArrayContent"> Byte array content. </param>
    /// <param name="excludeSheetsWithNoneOrOneRows"> If <c> True </c> exclude sheets with none or one rows, 
    /// if <c>False</c> the result could contain any <c>IEnumerable</c> with no <c>Dictionary</c> items. </param>
    /// <returns> <para>An <c> IEnumerable </c> where each item represents a sheet. 
    /// Each sheet contains a sequence of <c>Dictionary</c>, a single <c>Dictionary</c> represents only one row of the sheet.</para>
    /// The key of the <c>Dictionary</c> is the column name, and the value is the stored on the current field.</returns>
    public IEnumerable<IEnumerable<Dictionary<string, object?>>> ProcessExtractData(byte[] byteArrayContent, bool excludeSheetsWithNoneOrOneRows)
        => ValidateProcessExtractData(byteArrayContent, excludeSheetsWithNoneOrOneRows: excludeSheetsWithNoneOrOneRows);

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
    public IEnumerable<IEnumerable<Dictionary<string, object?>>> ProcessExtractData(byte[] byteArrayContent, IEnumerable<ExcelSheetField> fields, bool ignoreUnindicatedFields, bool excludeSheetsWithNoneOrOneRows)
        => ValidateProcessExtractData(byteArrayContent, fields, ignoreUnindicatedFields, excludeSheetsWithNoneOrOneRows);

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
    public IEnumerable<T> ProcessExtractDataSheet<T>(byte[] byteArrayContent, IEnumerable<ExcelField> fields, bool ignoreUnindicatedFields, int sheetIndex = 0)
        => ValidateProcessExtractDataSheet<T>(byteArrayContent, sheetIndex, fields, ignoreUnindicatedFields);

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
    public IEnumerable<T> ProcessExtractDataSheet<T>(byte[] byteArrayContent, bool ignoreUnindicatedFields, int sheetIndex = 0)
    {
        List<ExcelField> fields = GetFieldsFromModel<T>();

        return ValidateProcessExtractDataSheet<T>(byteArrayContent, sheetIndex, fields, ignoreUnindicatedFields);
    }

    private static List<ExcelField> GetFieldsFromModel<T>()
    {
        Type type = typeof(T);
        PropertyInfo[] fields = type.GetProperties();
        List<ExcelField> excelFields = new();

        foreach (PropertyInfo property in fields)
        {
            ExcelFieldAttribute? fieldAttribute = property.GetAttribute<ExcelFieldAttribute>();

            if (fieldAttribute is null)
                throw new MissingExcelFieldAttributeException();

            excelFields.Add(new ExcelField()
            {
                ColumnName = fieldAttribute.ColumnName,
                Required = fieldAttribute.Required,
                Type = fieldAttribute.Type
            });
        }

        return excelFields;
    }

    private IEnumerable<IEnumerable<Dictionary<string, object?>>> ValidateProcessExtractData(byte[] byteArrayContent, IEnumerable<ExcelSheetField>? fields = null, bool? ignoreUnindicatedFields = null, bool excludeSheetsWithNoneOrOneRows = false)
    {
        List<List<Dictionary<string, object?>>> excelData;

        using (MemoryStream stream = new(byteArrayContent, 0, byteArrayContent.Length))
        {
            Workbook workbook = GetWorkbook(stream);
            int sheetCountFile = workbook.Worksheets.Count;

            if (fields is not null && fields.Any(x => x.SheetIndex >= sheetCountFile))
                throw new SheetIndexNoExists();

            excelData = new(sheetCountFile);

            foreach (var worksheet in workbook.Worksheets.OrderBy(x => x.Index))
            {
                IEnumerable<ExcelSheetField>? sheetFields = fields?.Where(x => x.SheetIndex == worksheet.Index);
                int rowsNumber = worksheet.Cells.Rows.Count;
                bool hasNoneRow = rowsNumber == 0,
                    hasOneRow = rowsNumber == 1,
                    hasMoreThanOneRow = rowsNumber > 1;
                
                if (excludeSheetsWithNoneOrOneRows && (hasNoneRow || hasOneRow)) continue;

                List<Dictionary<string, object?>> excelDataSheet = new();
                
                if (hasMoreThanOneRow)
                    excelDataSheet = GetExcelDataSheet(worksheet, sheetFields, ignoreUnindicatedFields);
                    
                excelData.Add(excelDataSheet);
            }
        }

        bool existsOneSheetWithData = ExistsDataSheetNoEmpty(excelData.ToArray());

        if (!existsOneSheetWithData)
            throw new FileHasNoDataException();

        return excelData;
    }

    private IEnumerable<T> ValidateProcessExtractDataSheet<T>(byte[] byteArrayContent, int sheetIndex = 0, IEnumerable<ExcelField>? fields = null, bool? ignoreUnindicatedFields = null)
    {
        List<T> excelDataSheet = new();

        using (MemoryStream stream = new(byteArrayContent, 0, byteArrayContent.Length))
        {
            Workbook workbook = GetWorkbook(stream);

            Worksheet? worksheet = workbook.Worksheets.FirstOrDefault(x => x.Index == sheetIndex);

            if (worksheet is null)
                throw new SheetIndexNoExists();

            if (worksheet.Cells.Rows.Count == 0)
                throw new SheetHasNoRowException($"Sheet number {worksheet.Index + 1} has no rows.");

            IEnumerable<Dictionary<string, object?>> sheetData = GetExcelDataSheet(worksheet, fields, ignoreUnindicatedFields);

            bool dataSheetHasData = ExistsDataSheetNoEmpty(sheetData);

            if (!dataSheetHasData) 
                throw new SheetHasOnlyOneRowException($"Sheet number {worksheet.Index + 1} has only one row.");

            excelDataSheet = ConvertSheetData<T>(sheetData);
        }

        return excelDataSheet;
    }

    private static Workbook GetWorkbook(MemoryStream stream)
    {
        try
        {
            LoadOptions loadOptions = new(LoadFormat.Xlsx);
            Workbook workbook = new(stream, loadOptions);

            DeleteEmptyRows(workbook.Worksheets);

            return workbook;
        }
        catch
        {
            throw new UnsupportedFileException();
        }
    }

    private List<Dictionary<string, object?>> GetExcelDataSheet(Worksheet excelSheet, IEnumerable<IExcelField>? fields = null, bool? ignoreUnindicatedFields = null)
    {
        RowCollection rows = excelSheet.Cells.Rows;

        int columnsNumber = excelSheet.Cells.Count,
            rowsNumber = excelSheet.Cells.Rows.Count;

        Row columnsNameRow = rows[0];
        bool validateFields = fields?.Any() ?? false;

        if (validateFields)
            ValidateColumnsNames(columnsNameRow, columnsNumber, fields!, ignoreUnindicatedFields);

        int rowIndex = 1;
        List<Dictionary<string, object?>> excelDataSheet = new();

        while (rowIndex < rowsNumber)
        {
            Row row = rows[rowIndex];
            Dictionary<string, object?> rowData = GetRowData(row, columnsNameRow, columnsNumber, validateFields, fields, ignoreUnindicatedFields);
            excelDataSheet.Add(rowData);

            rowIndex++;
        }

        return excelDataSheet;
    }

    private static void DeleteEmptyRows(WorksheetCollection sheets)
    {
        for (int i = 0; i < sheets.Count; i++)
        {
            RowCollection rows = sheets[i].Cells.Rows;
            int rowsNumber = rows.Count;

            if (rowsNumber == 0) return;

            for (int rowIndex = rowsNumber - 1; rowIndex >= 0; rowIndex--)
            {
                Row row = rows.GetRowByIndex(rowIndex);

                bool emptyRow = true;

                foreach (Cell column in row)
                {
                    object? value = column.Type is CellValueType.IsNull ? null : column.Value;

                    if (!string.IsNullOrWhiteSpace(value?.ToString()))
                    {
                        emptyRow = false;
                        break;
                    }
                }

                if (!emptyRow) continue;

                rows.RemoveAt(rowIndex);
            }
        }    
    }

    private static void ValidateColumnsNames(Row row, int columnsNumber, IEnumerable<IExcelField> fields, bool? ignoreUnindicatedFields = null)
    {
        if (fields.Any(x => string.IsNullOrWhiteSpace(x.ColumnName)))
            throw new ExcelFieldColumnNameNullEmptyException();

        List<string> sheetColumnsNames = new();

        for (int i = 0; i < columnsNumber; i++)
        {
            string? columnName = row[i]?.Value?.ToString();

            if (string.IsNullOrEmpty(columnName)) continue;

            sheetColumnsNames.Add(columnName!.Trim());
        }

        IEnumerable<string> repeatedColumns = sheetColumnsNames.GroupBy(x => x).Where(x => x.Count() > 1).Select(x => x.Key);

        if (repeatedColumns.Any())
            throw new RepeatedColumnException($"Following columns are repeated: {string.Join(",", repeatedColumns)}");

        IEnumerable<string> columns = fields.Select(x => x.ColumnName);
        IEnumerable<string> missingColumns = columns.Where(x => !sheetColumnsNames.Any(y => y.Equals(x, StringComparison.OrdinalIgnoreCase)));

        if (missingColumns.Any())
            throw new MissingColumnException($"Missing columns: {string.Join(",", missingColumns)}");

        IEnumerable<string> noExistingColumns = sheetColumnsNames.Where(x => !columns.Any(y => y.Equals(x, StringComparison.OrdinalIgnoreCase)));

        if (noExistingColumns.Any() && ignoreUnindicatedFields.HasValue && !ignoreUnindicatedFields.Value)
            throw new NotIndicatedColumnNameException($"Following columns: {string.Join(",", noExistingColumns)} were not indicated as fields");
    }

    private Dictionary<string, object?> GetRowData(Row row, Row columnsNameRow, int columnsNumber, bool validateFields, IEnumerable<IExcelField>? fields = null, bool? ignoreUnindicatedFields = null)
    {
        int columnIndex = 0;
        Dictionary<string, object?> rowData = new();

        while (columnIndex < columnsNumber)
        {
            Cell column = row[columnIndex];
            string? columnName = columnsNameRow[columnIndex].Value?.ToString()?.Trim();
            bool columnHasName = !string.IsNullOrWhiteSpace(columnName);
            object? columnValue = column.Type is CellValueType.IsNull ? null : column.Value;
            bool columnHasValue = !string.IsNullOrWhiteSpace(columnValue?.ToString());

            if (!columnHasName && !columnHasValue)
            {
                columnIndex++;
                continue;
            }

            if (!columnHasName && columnHasValue)
                throw new FieldHasValueNoColumnNameException($"Column {columnsNameRow[columnIndex].Name} has no name, but has a value");

            if (validateFields && ignoreUnindicatedFields.HasValue && ignoreUnindicatedFields.Value && !fields!.Any(x => x.ColumnName.Equals(columnName!, StringComparison.OrdinalIgnoreCase)))
            {
                columnIndex++;
                continue;
            }

            if (validateFields)
            {
                IExcelField field = fields!.First(x => x.ColumnName.Equals(columnName!, StringComparison.OrdinalIgnoreCase));
                ValidateField(field, columnName!, columnHasValue, columnValue);
            }

            rowData.Add(columnName!, columnValue);

            columnIndex++;
        }
        return rowData;
    }

    private void ValidateField(IExcelField field, string columnName, bool columnHasValue, object? columnValue)
    {
        bool requiredFieldHasNoValue = field.Required && !columnHasValue;

        if (requiredFieldHasNoValue)
            throw new RequiredFieldException($"Required field '{columnName}' value is missing");

        bool validateDataType = field.Type is not null && columnHasValue;

        if (!validateDataType) return;

        bool existsDataType = Enum.IsDefined(typeof(DataTypes), field.Type!);

        if (!existsDataType)
            throw new ExcelFieldDataTypeNoExistsException($"Following data type {field.Type} of field {field.ColumnName} no exists");

        DataTypes dataType = EnumHelper.GetValues<DataTypes>().First(x => x == field.Type);

        bool matchType = _typeConverter.TryParse(columnValue!, dataType, out object? valueConverted);

        if (!matchType)
            throw new FieldValueTypeDifferentFieldDataTypeException($"Column {columnName} values must be of type {dataType.GetAttribute<DescriptionAttribute>()!.Description}");

        columnValue = valueConverted;
    }

    private static bool ExistsDataSheetNoEmpty(params IEnumerable<Dictionary<string, object?>>[] dataSheets) => dataSheets.Any(x => x.Any());

    private static List<T> ConvertSheetData<T>(IEnumerable<Dictionary<string, object?>> excelSheetData)
    {
        string jsonExcelData = JsonConvert.SerializeObject(excelSheetData);

        List<T> excelDataConverted = JArray.Parse(jsonExcelData).ToObject<List<T>>()!;

        return excelDataConverted;
    }
}