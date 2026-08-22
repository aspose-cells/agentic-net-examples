// Title: Count cells by data type in an Aspose.Cells worksheet using C# enumeration of CellValueType
// AI Prompts: Write C# code that uses Aspose.Cells to iterate over all cells in a worksheet and returns a dictionary mapping each CellValueType to its occurrence count. | Create a method that receives a Worksheet object and outputs separate totals for numeric, string, datetime, boolean, null, error, and unknown cells. | Extend the sample to also track and report the number of formula cells in addition to the standard CellValueType categories.
// Common Searches: how to count different cell types in an Excel file with Aspose.Cells C# | Aspose.Cells enumerate cells and get CellValueType statistics .NET | C# sample to tally numeric, string, date, boolean cells using Aspose.Cells | retrieve cell type distribution from a workbook using Aspose.Cells API
// Tags: enumerate cells CellValueType Aspose.Cells | count Excel cell data types C# | Aspose.Cells cell type aggregation | worksheet cell type statistics | C# tally numeric string datetime cells

using System;
using Aspose.Cells;
using System.Collections;

namespace AsposeCellsExamples
{
    // The example creates a workbook, fills six cells with numeric, string, DateTime, boolean, null, and error values, then iterates over every instantiated cell using a foreach loop. A switch on Cell.Type increments counters for each CellValueType (Unknown, Null, Numeric, DateTime, String, Bool, Error). The counts are printed and the workbook is saved.
    class CountCellsByDataType
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data of various types
                cells["A1"].PutValue(123);                     // Numeric
                cells["A2"].PutValue("Hello World");           // String
                cells["A3"].PutValue(DateTime.Now);            // DateTime
                cells["A4"].PutValue(true);                    // Boolean
                cells["A5"].PutValue(null);                    // Null
                cells["A6"].PutValue("=1/0");                  // Error (division by zero)

                // Counters for each CellValueType
                int unknownCount = 0;
                int nullCount = 0;
                int numericCount = 0;
                int dateTimeCount = 0;
                int stringCount = 0;
                int boolCount = 0;
                int errorCount = 0;

                // Iterate through all instantiated cells using foreach
                foreach (Cell cell in cells)
                {
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
                Console.WriteLine($"Unknown : {unknownCount}");
                Console.WriteLine($"Null    : {nullCount}");
                Console.WriteLine($"Numeric : {numericCount}");
                Console.WriteLine($"DateTime: {dateTimeCount}");
                Console.WriteLine($"String  : {stringCount}");
                Console.WriteLine($"Bool    : {boolCount}");
                Console.WriteLine($"Error   : {errorCount}");

                // Save the workbook (optional, demonstrates lifecycle usage)
                string outputPath = "CountCellsByDataType.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
