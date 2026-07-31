// Title: Aspose.Cells for .NET – Log a warning when MaxDataRow = 0 and MaxDataColumn > 0
// Description: Demonstrates how to create a workbook, add data only to the first row across multiple columns, read the Cells.MaxDataRow and Cells.MaxDataColumn properties, and output a console warning if MaxDataRow is zero while MaxDataColumn is greater than zero, then save the file.
// Keywords: Aspose.Cells MaxDataRow | MaxDataColumn check | worksheet warning .NET | detect empty rows with data columns | Excel data range validation | log warning Aspose.Cells | C# worksheet data check
// Common Searches: Aspose.Cells log warning when MaxDataRow is zero | check MaxDataColumn greater than zero in .NET | validate worksheet data range Aspose.Cells | detect header‑only rows using MaxDataRow | C# Aspose.Cells MaxDataRow MaxDataColumn example
// Developer Intent: Emit a warning for worksheets that have no data rows but contain one or more populated columns.
// Use Cases: Flag Excel sheets that appear empty because only header columns are present, preventing downstream processing errors. | Validate imported workbooks before calculations by ensuring both row and column data exist. | Generate logs for worksheets with header rows only, helping data quality audits.
// AI Prompts: Create a method that scans all worksheets in a workbook and logs a warning for any sheet where MaxDataRow == 0 and MaxDataColumn > 0 using Aspose.Cells. | Show how to throw a custom exception instead of a console warning when MaxDataRow is zero but MaxDataColumn exceeds zero. | Integrate the MaxDataRow/MaxDataColumn check into a multi‑sheet validation routine that aggregates warnings and writes them to a log file.

using System;
using Aspose.Cells;

namespace AsposeCellsWarningDemo
{
    // Demonstrates how to create a workbook, add data only to the first row across multiple columns, read the Cells.MaxDataRow and Cells.MaxDataColumn properties, and output a console warning if MaxDataRow is zero while MaxDataColumn is greater than zero, then save the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Add data only to the first row but across multiple columns
            cells["A1"].PutValue("FirstColumn");
            cells["B1"].PutValue("SecondColumn"); // This makes MaxDataColumn > 0

            // Check the condition: MaxDataRow == 0 && MaxDataColumn > 0
            int maxDataRow = cells.MaxDataRow;       // Zero‑based index of the last row containing data
            int maxDataColumn = cells.MaxDataColumn; // Zero‑based index of the last column containing data

            if (maxDataRow == 0 && maxDataColumn > 0)
            {
                // Log a warning
                Console.WriteLine("Warning: Worksheet '{0}' has MaxDataRow = 0 but MaxDataColumn = {1}.",
                                  worksheet.Name, maxDataColumn);
            }

            // Save the workbook (adjust the path as needed)
            workbook.Save("WarningDemo.xlsx");
        }
    }
}
