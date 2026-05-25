using System;
using System.Collections;
using Aspose.Cells;

class CountCellTypes
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data with different value types
        cells["A1"].PutValue(123);                 // Numeric
        cells["A2"].PutValue("Hello World");       // String
        cells["A3"].PutValue(DateTime.Now);        // DateTime
        cells["A4"].PutValue(true);                // Boolean
        cells["A5"].PutValue(null);                // Null (blank)
        cells["A6"].Formula = "=1/0";              // Error (division by zero)

        // Initialize counters for each CellValueType
        int numericCount = 0;
        int stringCount = 0;
        int dateTimeCount = 0;
        int boolCount = 0;
        int nullCount = 0;
        int errorCount = 0;
        int unknownCount = 0;

        // Iterate through all instantiated cells using the enumerator
        IEnumerator enumerator = cells.GetEnumerator();
        while (enumerator.MoveNext())
        {
            Cell cell = (Cell)enumerator.Current;
            switch (cell.Type)
            {
                case CellValueType.IsNumeric:
                    numericCount++;
                    break;
                case CellValueType.IsString:
                    stringCount++;
                    break;
                case CellValueType.IsDateTime:
                    dateTimeCount++;
                    break;
                case CellValueType.IsBool:
                    boolCount++;
                    break;
                case CellValueType.IsNull:
                    nullCount++;
                    break;
                case CellValueType.IsError:
                    errorCount++;
                    break;
                default:
                    unknownCount++;
                    break;
            }
        }

        // Display the counts for each data type
        Console.WriteLine($"Numeric cells: {numericCount}");
        Console.WriteLine($"String cells: {stringCount}");
        Console.WriteLine($"DateTime cells: {dateTimeCount}");
        Console.WriteLine($"Boolean cells: {boolCount}");
        Console.WriteLine($"Null cells: {nullCount}");
        Console.WriteLine($"Error cells: {errorCount}");
        Console.WriteLine($"Unknown cells: {unknownCount}");

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("CellTypeCounts.xlsx");
    }
}