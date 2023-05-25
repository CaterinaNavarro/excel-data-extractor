# Excel Data Reader Extractor C# .NET 6 

Open Source library specialized in processing, validating and extracting data from files in Excel format. Supports field validations and conversion to a specific type.



## Installation

Available on **Nuget**

**https://www.nuget.org/packages/ExcelDataReaderExtractor/**



## Documentation

You can do DI using the interface

```csharp
IExcelDataReaderExtractor
```

which is implemented by

```csharp
ExcelDataReaderExtractor
```


The library extract the sheet data into the generic form
```csharp
IEnumerable<IEnumerable<Dictionary<string, object?>>>
```
Each item represents a sheet. Each sheet contains a sequence of Dictionary, a single Dictionary represents only one row of the sheet.
The key of the Dictionary is the column name, and the value is the stored on the current field.

It provides methods to convert each Dictionary element into a specific object T type, to do this is necessary the properties that this T type has, contains JsonPropertyAttribute (or similar, if necessary) or ExcelFieldAttribute for matching with columns names that are stored as keys of the dictionary.

Newtonsoft.Json is used to convert the objects.


#### Extract all data of all sheets

```csharp
IEnumerable<IEnumerable<Dictionary<string, object?>>> ProcessExtractData(byte[] byteArrayContent, bool excludeSheetsWithNoneOrOneRows);
```

Params:
* byteArrayContent
* excludeSheetsWithNoneOrOneRows: If True exclude sheets with none or one rows, if False the result could contain any IEnumerable with no Dictionary items.

Returns An IEnumerable where each item represents a sheet. Each sheet contains a sequence of Dictionary, a single Dictionary represents only one row of the sheet.
The key of the Dictionary is the column name, and the value is the stored on the current field.

##### ---
#### Extract specific data, performs fields validations

```csharp
IEnumerable<IEnumerable<Dictionary<string, object?>>> ProcessExtractData(byte[] byteArrayContent, IEnumerable<ExcelSheetField> fields, bool ignoreUnindicatedFields, bool excludeSheetsWithNoneOrOneRows);
```

Params:
* byteArrayContent
* fields: Fields that the sheets must contain.
* ignoreUnindicatedFields: If true does not make any validations on the fields that exists in the file but were not indicated as fields, as consequence it does not extract them neither.
* excludeSheetsWithNoneOrOneRows: If True exclude sheets with none or one rows, if False the result could contain any IEnumerable with no Dictionary items.

Returns an IEnumerable where each item represents a sheet. 
Each sheet contains a sequence of Dictionary, a single Dictionary represents only one row of the sheet.
The key of the Dictionary is the column name, and the value is the stored on the current field.

##### ---
#### Extract the data of a specific sheet, performing fields validations

```csharp
IEnumerable<T> ProcessExtractDataSheet<T>(byte[] byteArrayContent, IEnumerable<ExcelField> fields, bool ignoreUnindicatedFields, int sheetIndex = 0);
```

Params:
* T: Output class whose properties contains JsonPropertyAttribute (or another, if necessary) for matching the columns names.
* byteArrayContent: Byte array content.
* fields: Fields that the sheet must contain.
* ignoreUnindicatedFields: If true does not make any validations on the fields that exists in the sheet but were not indicated as fields, as consequence it does not extract them neither. If false validate the sheet contains the columns indicated only.
* sheetIndex: Sheet index to extract, as default is the first. 

Returns an IEnumerable containing the rows converted into the output type.


##### ---
#### Extract the data of a specific sheet, performing fields validations. Properties of the output type T have ExcelFieldAttribute

```csharp
List<T> ProcessExtractDataSheet<T>(byte[] byteArrayContent, bool ignoreUnindicatedFields, int sheetIndex = 0);
```

Params: 
* T: Output class whose properties contains the ExcelFieldAttribute for matching the columns names and provide specific information of the fields.
* byteArrayContent
* ignoreUnindicatedFields: If true does not make any validations on the fields that exists in the sheet but were not indicated as fields, as consequence it does not extract them neither.
If false validate the sheet contains the columns indicated only.
* sheetIndex: Sheet index to extract, as default it is the first. 

Returns an IEnumerable containing the rows converted into the output type.



## Validations
* File content must be in byte array form
* Support for XLSX format files
* When using ProcessExtractDataSheet<T> the sheet must contain at least one row besides the columns names row
* When using ProcessExtractData empty or only one row sheets are not validated (not throw the related exceptions), so the result can contain 
sequences of sheets with no Dictionary items if excludeSheetsWithNoneOrOneRows is false, if true the result only contains sequences of sheets 
that have at least two rows
* Fields with value and no column name are not valid
* Optional data type field validation for integers and strings



## Usage

You can find examples on the test project

#### Return data without converting into a specific type
```csharp
IExcelDataReaderExtractor _excelDataReaderExtractor = new ExcelDataReaderExtractor();

public void Extract_All_Data_No_Convert_Model()
{
    IEnumerable<IEnumerable<Dictionary<string, object?>>> excelData;

    excelData = _excelDataReaderExtractor.ProcessExtractData(_thirdSheetHasValuesContent, excludeSheetsWithNoneOrOneRows: false);

    Assert.True(excelData.Count() == 3 && excelData.Last().Count() == 1);
}

public void Extract_Data_Excluding_Sheets_With_None_One_Row()
{
    IEnumerable<IEnumerable<Dictionary<string, object?>>> excelData;

    excelData = _excelDataReaderExtractor.ProcessExtractData(_thirdSheetHasValuesContent, excludeSheetsWithNoneOrOneRows: true);

    Assert.True(excelData.Count() == 1 && excelData.First().Count() == 1);
    
}

public void Extract_Data_Validate_Fields_No_Convert_Model()
{
    IEnumerable<IEnumerable<Dictionary<string, object?>>> excelData;
    List<ExcelSheetField> fields = new()
    {
        new()
        {
            ColumnName = "FirstColumnNumber",
            Required = true,
            Type = DataTypes.Integer,
            SheetIndex = 0,
        },
        new()
        {
            ColumnName = "SecondColumnStringSecondSheet",
            Required = true,
            Type= DataTypes.String,
            SheetIndex = 1
        }
    };

    int firstColumnFirstSheetValue = 1;
    string secondColumnSecondSheetValue = "fifth value";

    excelData = _excelDataReaderExtractor.ProcessExtractData(_dataOnTwoSheetsContent, fields, ignoreUnindicatedFields: true, excludeSheetsWithNoneOrOneRows: false);

    Assert.True(excelData.Count() == 2 && 
        excelData.First().Any(firstSheet => (int)firstSheet["FirstColumnNumber"]! == firstColumnFirstSheetValue) &&
        excelData.Last().Any(secondSheet => secondSheet["SecondColumnStringSecondSheet"]!.ToString() == secondColumnSecondSheetValue));
}
```

#### Return data converting into a specific type

Output type T examples

```csharp
internal class ExcelDataRow
{
    [JsonProperty("FirstColumnNumber")]
    public int FirstColumn { get; set; }

    [JsonProperty("SecondColumnString")]
    public string SecondColumn { get; set; } = null!;
}

internal class ExcelDataRowWithFieldAttribute
{
    [ExcelField(columnName: "FirstColumnNumber", required: true, type: DataTypes.Integer)]
    public int FirstColumn { get; set; }
    
    [ExcelField(columnName: "SecondColumnString", required: true, type: DataTypes.String)]
    public string SecondColumn { get; set; } = null!;
}
```

```csharp
public void Extract_Data_Sheet_Fields_Convert_Model()
{
    IEnumerable<ExcelDataRow> excelDataSheet;
    List<ExcelField> fields = new()
    {
        new()
        {
            ColumnName = "FirstColumnNumber",
            Required = true,
            Type = DataTypes.Integer
        },
        new ()
        {
            ColumnName = "SecondColumnString",
            Required = true,
            Type = DataTypes.String
        }
    };

    excelDataSheet = _excelDataReaderExtractor.ProcessExtractDataSheet<ExcelDataRow>(_columnsWithDataContent, fields: fields, ignoreUnindicatedFields: true);

    Assert.True(excelDataSheet.Count() == 2);
}

public void Extract_Data_Second_Sheet_Fields_Convert_Model()
{
    IEnumerable<ExcelDataRowSecondSheet> excelDataSheet;
    List<ExcelField> fields = new()
    {
        new()
        {
            ColumnName = "FirstColumnNumberSecondSheet",
            Required = true,
            Type = DataTypes.Integer
        },
        new()
        {
            ColumnName = "SecondColumnStringSecondSheet",
            Required = false,
            Type = DataTypes.String
        }
    };

    int firstColumnValue = 5;
    string secondColumnValue = "fifth value";

    excelDataSheet = _excelDataReaderExtractor.ProcessExtractDataSheet<ExcelDataRowSecondSheet>(_dataOnTwoSheetsContent, fields: fields, ignoreUnindicatedFields: false, sheetIndex: 1);

    Assert.Contains(excelDataSheet, x => x.FirstColumn == firstColumnValue && x.SecondColumn == secondColumnValue);
}
        
public void Extract_Data_Sheet_Convert_Model_With_Fields_Attribute()
{
    IEnumerable<ExcelDataRowWithFieldAttribute> excelDataSheet;

    excelDataSheet = _excelDataReaderExtractor.ProcessExtractDataSheet<ExcelDataRowWithFieldAttribute>(_columnsWithDataContent, ignoreUnindicatedFields: true);

    Assert.NotEmpty(excelDataSheet);
}
        
```


## Exceptions

### Column Exceptions
Inherits from ColumnException

#### MissingColumnException

Exception thrown when the sheet does not contains all columns names given by ExcelField, ExcelSheetField or ExcelFieldAttribute.

#### NotIndicatedColumnNameException

Exception thrown when ignoreUnindicatedFields is false, the process validate all the columns names of the sheet  and the sheet has more columns names than the given by ExcelField, ExcelSheetField or ExcelFieldAttribute.

#### RepeatedColumnException

Exception thrown when a sheet as repeated columns names.


##### ---
### Field Exceptions
Inherits from FieldException

#### ExcelFieldColumnNameNullEmptyException

Exception thrown when the given field ExcelField, ExcelSheetField or ExcelFieldAttribute has no column name.

#### ExcelFieldDataTypeNoExistsException

Exception thrown when given field ExcelField, ExcelSheetField or ExcelFieldAttribute has DataTypes value not existing in the enum.

#### FieldHasValueNoColumnNameException

Exception thrown when a field has value but its column name is missing.

#### FieldValueTypeDifferentFieldDataTypeException

Exception thrown when the field type value is different from the one given by DataTypes value.

#### MissingExcelFieldAttributeException

Exception thrown when using the method to extract data and convert into a specific object type T without indicating the list of fields apart.
In this case, the type T must include the ExcelFieldAttribute in all of its properties.

#### RequiredFieldException

Exception thrown when a field is required and has no value.


##### ---
### Sheet Exceptions
Inherits from SheetException

#### SheetHasNoRowException
Exception thrown when the sheet has no rows.

#### SheetHasOnlyOneRowException
Exception thrown when a sheet has only one row.

#### SheetIndexNoExists
Exception thrown when the sheet index provided does not exists in the file.


#### ---
### File Exceptions
Inherits from FileException

#### UnsupportedFileException

Exception thrown when trying to process an unsupported file.

#### FileHasNoDataException
Exception thrown when the sheet/s of the file are empty or have only one row.