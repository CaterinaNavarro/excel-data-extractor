# Excel Data Reader Extractor C# .NET 6 

Nuget library specialized in process, validate and extract excel file data.


## Documentation

The library extract the sheet data into the generic form
```csharp
List<List<Dictionary<string, object?>>>
```
Each list item of the main list represents a sheet.
Each sheet contains a list of dictionary, a dictionary represents only one row of the sheet.
The key of the dictionary is the column name, and the value is the stored on the current field.

It provides methods to parse each Dictionary element into an specific object T type, to do this is necessary the properties that this T type has, contains JsonPropertyAttribute (or similar, if necessary) or ExcelFieldAttribute for matching with columns names that are stored as keys of the dictionary.

Newtonsoft.Json is used to convert the objects.


#### Extract all data of all sheets

```csharp
List<List<Dictionary<string, object?>>> ProcessExtractData(byte[] byteArrayContent);
```

Params:
* byteArrayContent
Returns List of List of Dictionary, each list item of the main list represents a sheet.
Each sheet contains a list of dictionary, a dictionary represents only one row of the sheet.
The key of the dictionary is the column name, and the value is the stored on the current field.

##### ---
#### Extract specific data, performs fields validations

```csharp
List<List<Dictionary<string, object?>>> ProcessExtractData(byte[] byteArrayContent, IEnumerable<ExcelSheetField> fields, bool ignoreUnindicatedFields);
```

Params:
* byteArrayContent
* fields: Fields that the sheets must contain.
* ignoreUnindicatedFields: If true does not make any validations on the fields that exists in the file but were not indicated as fields, as consequence it does not extract them neither. If false validate the sheets contains the columns indicated only.

Returns a List of List of Dictionary, each list item of the main list represents a sheet. Each sheet contains a list of dictionary, a dictionary represents only one row of the sheet. The key of the dictionary is the column name, and the value is the stored on the current field.

##### ---
#### Extract the data of a specific sheet

```csharp
List<T> ProcessExtractDataSheet<T>(byte[] byteArrayContent, int sheetIndex = 0);
```

Params:
* T: Output class whose properties contains JsonPropertyAttribute (or similar, if necessary) for matching the columns names.
* byteArrayContent
* sheetIndex: Sheet index to extract, as default it is the first.

Returns the rows parsed into the output class list. 


##### ---
#### Extract the data of a specific sheet, performing fields validations

```csharp
List<T> ProcessExtractDataSheet<T>(byte[] byteArrayContent, IEnumerable<ExcelField> fields, bool ignoreUnindicatedFields, int sheetIndex = 0);
```

Params:
* T: Output class whose properties contains JsonPropertyAttribute (or similar, if necessary) for matching the columns names.
* byteArrayContent: Byte array content.
* fields: Fields that the sheet must contain.
* ignoreUnindicatedFields: If true does not make any validations on the fields that exists in the sheet but were not indicated as fields,as consequence it does not extract them neither. If false validate the sheet contains the columns indicated only.
* sheetIndex: Sheet index to extract, as default is the first. 

Returns the rows parsed into the output class list.


##### ---
#### Extract the data of a specific sheet, performing fields validations. Properties of the output type T have ExcelFieldAttribute

Params: 
* T: Output class whose properties contains the ExcelFieldAttribute for matching the columns names and provide specific information of the fields.
* byteArrayContent
* ignoreUnindicatedFields: If true does not make any validations on the fields that exists in the sheet but were not indicated as fields, as consequence it does not extract them neither. If false validate the sheet contains the columns indicated only.
* sheetIndex: Sheet index to extract, as default it is the first. 
* The rows parsed into the output class list. 

```csharp
List<T> ProcessExtractDataSheet<T>(byte[] byteArrayContent, bool ignoreUnindicatedFields, int sheetIndex = 0);
```


## Validations
* File content must be in byte array form
* Support for XLSX format files
* Sheet must not be empty
* First row of every sheet must contain columns names
* Any sheet must contain at least one row besides the columns names row
* Fields with value and no column name are not valid
* Data type field validation for integers and strings
## Usage


```csharp
public void Extract_All_Data_No_Parse_Model()
{
    List<List<Dictionary<string, object?>>> excelData;

    excelData = _excelDataReaderExtractor.ProcessExtractData(_columnsWithDataContent);

    Assert.NotEmpty(excelData);
}


public void Extract_Data_Validate_Fields_No_Parse_Model()
{
    List<List<Dictionary<string, object?>>> excelData;
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

    excelData = _excelDataReaderExtractor.ProcessExtractData(_dataOnTwoSheetsContent, fields, ignoreUnindicatedFields: true);

    Assert.True(excelData[0].Any(firstSheet => (int)firstSheet["FirstColumnNumber"]! == firstColumnFirstSheetValue) &&
                excelData[1].Any(secondSheet => secondSheet["SecondColumnStringSecondSheet"]!.ToString() == secondColumnSecondSheetValue));
}


public void Extract_Data_Sheet_Parse_Model()
{
    List<ExcelDataRow> excelDataSheet;

    excelDataSheet = _excelDataReaderExtractor.ProcessExtractDataSheet<ExcelDataRow>(_columnsWithDataContent);
        
    Assert.NotEmpty(excelDataSheet);
}


public void Extract_Data_Sheet_Fields_Parse_Model()
{
    List<ExcelDataRow> excelDataSheet;
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

    Assert.NotEmpty(excelDataSheet);
}


public void Extract_Data_Second_Sheet_Fields_Parse_Model()
{
    List<ExcelDataRowSecondSheet> excelDataSheet;
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


public void Extract_Data_Sheet_Parse_Model_With_Fields_Attribute()
{
    List<ExcelDataRowWithFieldAttribute> excelDataSheet;

    excelDataSheet = _excelDataReaderExtractor.ProcessExtractDataSheet<ExcelDataRowWithFieldAttribute>(_columnsWithDataContent, ignoreUnindicatedFields: true);

    Assert.NotEmpty(excelDataSheet);
}
```
More examples on the testing project.

## Exceptions

### Column Exceptions
Inherits from ColumnException

#### MissingColumnException

Exception thrown when the sheet does not contains all columns names given by ExcelField, ExcelSheetField or ExcelFieldAttribute.

#### MissingColumnNameFirstRowException

Exception thrown when a sheet does not contain any column name in the first row.

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

Exception thrown when using the method to extract data and parse into an specific object type T without indicating the list of fields apart. In this case, the type T must include the ExcelFieldAttribute in all of its properties.

#### RequiredFieldException

Exception thrown when a field is required and has no value.


##### ---
### Sheet Exceptions
Inherits from SheetException

#### EmptySheetException

Exception thrown when the sheet it is empty.

#### SheetIndexNoExists

Exception thrown when the sheet index provided does not exists in the file.

#### ---
### File Exceptions

#### UnsupportedFileException

Exception thrown when trying to process an unsupported file.