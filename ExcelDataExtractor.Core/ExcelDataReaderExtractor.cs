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

public class ExcelDataReaderExtractor : IExcelDataReaderExtractor
{
    private readonly TypeConverterHelper _typeConverter = new();

    public List<List<Dictionary<string, object?>>> ProcessExtractData(byte[] byteArrayContent)
        => ValidateProcessExtractData(byteArrayContent);

    public List<List<Dictionary<string, object?>>> ProcessExtractData(byte[] byteArrayContent, IEnumerable<ExcelSheetField> fields, bool ignoreUnindicatedFields)
        => ValidateProcessExtractData(byteArrayContent, fields, ignoreUnindicatedFields);

    public List<T> ProcessExtractDataSheet<T>(byte[] byteArrayContent, IEnumerable<ExcelField> fields, bool ignoreUnindicatedFields, int sheetIndex = 0)
        => ValidateProcessExtractDataSheet<T>(byteArrayContent, sheetIndex, fields, ignoreUnindicatedFields);

    public List<T> ProcessExtractDataSheet<T>(byte[] byteArrayContent, bool ignoreUnindicatedFields, int sheetIndex = 0)
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

    private List<List<Dictionary<string, object?>>> ValidateProcessExtractData(byte[] byteArrayContent, IEnumerable<ExcelSheetField>? fields = null, bool? ignoreUnindicatedFields = null)
    {
        List<List<Dictionary<string, object?>>> excelData;

        using (MemoryStream stream = new(byteArrayContent, 0, byteArrayContent.Length))
        {
            Workbook workbook = GetWorkbook(stream);
            int sheetCountFile = workbook.Worksheets.Count;

            if (fields is not null && fields.Any(x => x.SheetIndex >= sheetCountFile))
                throw new SheetIndexNoExists();

            excelData = new(sheetCountFile);

            foreach (var worksheet in workbook.Worksheets)
            {
                IEnumerable<ExcelSheetField>? sheetFields = fields?.Where(x => x.SheetIndex == worksheet.Index);
                List<Dictionary<string, object?>> excelDataSheet = GetExcelDataSheet(worksheet, sheetFields, ignoreUnindicatedFields);
                excelData.Add(excelDataSheet);
            }
        }

        return excelData;
    }

    private List<T> ValidateProcessExtractDataSheet<T>(byte[] byteArrayContent, int sheetIndex = 0, IEnumerable<ExcelField>? fields = null, bool? ignoreUnindicatedFields = null)
    {
        List<T> excelDataSheet = new();

        using (MemoryStream stream = new(byteArrayContent, 0, byteArrayContent.Length))
        {
            Workbook workbook = GetWorkbook(stream);

            Worksheet? worksheet = workbook.Worksheets.FirstOrDefault(x => x.Index == sheetIndex);

            if (worksheet is null)
                throw new SheetIndexNoExists();

            List<Dictionary<string, object?>> sheetData = GetExcelDataSheet(worksheet, fields, ignoreUnindicatedFields);

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

            return workbook;
        }
        catch
        {
            throw new UnsupportedFileException();
        }
    }

    private List<Dictionary<string, object?>> GetExcelDataSheet(Worksheet excelSheet, IEnumerable<IExcelField>? fields = null, bool? ignoreUnindicatedFields = null)
    {
        DeleteEmptyRows(excelSheet);

        RowCollection rows = excelSheet.Cells.Rows;

        int columnsNumber = excelSheet.Cells.Count,
            rowsNumber = excelSheet.Cells.Rows.Count;

        if (rowsNumber == 0)
            throw new EmptySheetException($"Sheet number {excelSheet.Index + 1} is empty");

        if (rowsNumber == 1)
            throw new MissingColumnNameFirstRowException();

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

    private static void DeleteEmptyRows(Worksheet excelSheet)
    {
        RowCollection rows = excelSheet.Cells.Rows;
        int rowsNumber = rows.Count;

        if (rowsNumber == 0) return;

        for (int rowIndex =  - 1; rowIndex >= 0; rowIndex--)
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

        bool validateDataType = field.Type is not null;

        if (!validateDataType) return;

        bool existsDataType = Enum.IsDefined(typeof(DataTypes), field.Type!);

        if (!existsDataType)
            throw new ExcelFieldDataTypeNoExistsException($"Following data type {field.Type} of field {field.ColumnName} no exists");

        DataTypes dataType = EnumHelper.GetValues<DataTypes>().First(x => x == field.Type);

        if (field.Required)
        {
            bool matchType = _typeConverter.TryParse(columnValue!, dataType, out object? valueConverted);

            if (!matchType)
                throw new FieldValueTypeDifferentFieldDataTypeException($"Column {columnName} values must be of type {dataType.GetAttribute<DescriptionAttribute>()!.Description}");

            columnValue = valueConverted;
        }
    }

    private static List<T> ConvertSheetData<T>(List<Dictionary<string, object?>> excelSheetData)
    {
        string jsonExcelData = JsonConvert.SerializeObject(excelSheetData);

        List<T> excelDataConverted = JArray.Parse(jsonExcelData).ToObject<List<T>>()!;

        return excelDataConverted;
    }
}