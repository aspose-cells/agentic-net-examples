// Title: How to add exception handling for invalid TableCssId when styling a ListObject table with Aspose.Cells in C#
// AI Prompts: Write C# code that assigns a TableCssId to a ListObject in Aspose.Cells, wrapping the assignment in a try‑catch block that catches CellsException and logs a clear message if the ID is invalid. | Show how to verify that a TableCssId string conforms to CSS identifier rules before setting it on a ListObject, and throw an ArgumentException with a descriptive error when it does not. | Provide a modified version of the workbook‑creation example that includes custom TableCssId handling, catches both CellsException and generic Exception, and outputs user‑friendly error information.
// Common Searches: Aspose.Cells C# catch CellsException for invalid TableCssId | validate TableCssId before applying to ListObject Aspose.Cells | how to handle invalid CSS ID in Aspose.Cells table styling | C# example of error handling when setting TableCssId in Aspose.Cells | Aspose.Cells ListObject custom CSS ID error message
// Tags: Aspose.Cells ListObject custom style ID handling | C# error handling for table style identifiers | validate table CSS identifier Aspose.Cells | catch CellsException for invalid style ID | Aspose.Cells workbook creation error handling

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables; // Required for ListObject and TableStyleType

namespace AsposeCellsTableCssIdExample
{
    // The example creates a workbook, adds a ListObject table with sample data, applies a built‑in table style, and demonstrates how to safely set a custom TableCssId. It includes try‑catch blocks for CellsException and generic Exception, providing clear console messages that help developers diagnose invalid CSS identifiers or other issues during table styling.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the output Excel file
            string outputPath = "TableWithCssId.xlsx";

            try
            {
                // Ensure the output directory exists (if a directory is specified)
                string? outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Create a new workbook and access the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some data for the table
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("Alice");
                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue("Bob");

                // Define the range for the table (A1:B3)
                int firstRow = 0;      // Zero‑based index
                int firstColumn = 0;   // Zero‑based index
                int totalRows = 3;     // Number of rows in the range
                int totalColumns = 2;  // Number of columns in the range
                bool hasHeaders = true;

                // Add the ListObject (table) and retrieve the created object
                int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, hasHeaders);
                ListObject table = sheet.ListObjects[tableIndex];

                // Apply a built‑in table style (choose any available style)
                table.TableStyleType = TableStyleType.TableStyleMedium2;

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (CellsException ex)
            {
                // Handle Aspose.Cells specific exceptions
                Console.Error.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Catch any other unexpected exceptions
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
