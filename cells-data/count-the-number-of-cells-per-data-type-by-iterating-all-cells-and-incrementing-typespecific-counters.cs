// Title: C# – Count Cells by Data Type in an Aspose.Cells Worksheet
// Description: Creates a workbook, populates cells with numeric, string, DateTime, Boolean, null and error values, then enumerates the worksheet's Cells collection. For each cell the Cell.Type (CellValueType) is examined and counters for Unknown, Null, Numeric, DateTime, String, Boolean and Error are incremented. The totals are printed and the workbook saved.
// Keywords: Aspose.Cells C# count cells by type | CellValueType enumeration | worksheet cell data type statistics | iterate Aspose.Cells cells | Excel cell type detection Aspose | C# Aspose.Cells cell counters | enumerate cells Aspose.Cells API
// Common Searches: how to count cells by data type using Aspose.Cells C# | enumerate worksheet cells and get CellValueType | Aspose.Cells count numeric and string cells | C# get cell type distribution in Excel workbook | Aspose.Cells tally error cells in a sheet
// Developer Intent: The developer needs to determine how many cells of each data type exist in a worksheet and produce a summary of those counts.
// Use Cases: Generate a data‑type breakdown report to validate spreadsheet content before further processing. | Identify unexpected error or null cells for data cleansing or quality checks. | Calculate ratios of numeric versus textual entries for analytics on imported Excel data.
// AI Prompts: Write C# code using Aspose.Cells that returns a dictionary of CellValueType counts for a given worksheet. | Extend the example to separate formula cells from their evaluated values while counting types. | Suggest alternative Aspose.Cells methods that provide cell type statistics without manual enumeration.

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, populates cells with numeric, string, DateTime, Boolean, null and error values, then enumerates the worksheet's Cells collection. For each cell the Cell.Type (CellValueType) is examined and counters for Unknown, Null, Numeric, DateTime, String, Boolean and Error are incremented. The totals are printed and the workbook saved.
    class CountCellsByDataType
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data of various types
            cells["A1"].PutValue(123);                     // Numeric
            cells["A2"].PutValue("Hello World");           // String
            cells["A3"].PutValue(DateTime.Now);            // DateTime
            cells["A4"].PutValue(true);                    // Boolean
            cells["A5"].PutValue(null);                    // Null
            cells["A6"].PutValue("=1/0");                  // Error (division by zero)

            // Initialize counters for each CellValueType
            int unknownCount = 0;
            int nullCount = 0;
            int numericCount = 0;
            int dateTimeCount = 0;
            int stringCount = 0;
            int boolCount = 0;
            int errorCount = 0;

            // Iterate through all instantiated cells using the enumerator
            IEnumerator enumerator = cells.GetEnumerator();
            while (enumerator.MoveNext())
            {
                Cell cell = (Cell)enumerator.Current;

                // Increment the appropriate counter based on the cell's type
                switch (cell.Type)
                {
                    case CellValueType.IsUnknown:
                        unknownCount++;
                        break;
                    case CellValueType.IsNull:
                        nullCount++;
                        break;
                    case CellValueType.IsNumeric:
                        numericCount++;
                        break;
                    case CellValueType.IsDateTime:
                        dateTimeCount++;
                        break;
                    case CellValueType.IsString:
                        stringCount++;
                        break;
                    case CellValueType.IsBool:
                        boolCount++;
                        break;
                    case CellValueType.IsError:
                        errorCount++;
                        break;
                }
            }

            // Output the results
            Console.WriteLine("Cell counts by data type:");
            Console.WriteLine($"Unknown   : {unknownCount}");
            Console.WriteLine($"Null      : {nullCount}");
            Console.WriteLine($"Numeric   : {numericCount}");
            Console.WriteLine($"DateTime  : {dateTimeCount}");
            Console.WriteLine($"String    : {stringCount}");
            Console.WriteLine($"Boolean   : {boolCount}");
            Console.WriteLine($"Error     : {errorCount}");

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("CountCellsByDataType.xlsx");
        }
    }
}
