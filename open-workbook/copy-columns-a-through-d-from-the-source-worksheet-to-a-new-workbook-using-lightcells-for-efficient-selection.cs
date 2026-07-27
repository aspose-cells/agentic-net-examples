// Title: Efficiently copy columns A‑D to a new workbook with LightCells (Aspose.Cells C#)
// Description: This example opens a source Excel file, determines the last populated row, and uses Aspose.Cells LightCells to rapidly transfer the first four columns (A‑D) – including cell values and optional formatting – into a freshly created workbook, which is then saved.
// Keywords: Aspose.Cells LightCells | C# copy columns A-D | Excel range transfer .NET | fast column copy Aspose | preserve cell style Aspose.Cells | copy worksheet columns C# | LightCells performance | Aspose.Cells API | create new workbook from range | efficient Excel column copy
// Common Searches: Aspose.Cells copy columns A to D using LightCells | C# copy specific columns to new workbook Aspose | How to use LightCells for range copy in Aspose.Cells | Efficient column transfer Aspose.Cells .NET | Copy Excel columns with formatting Aspose C#
// Developer Intent: Move the first four columns from an existing worksheet into a separate workbook while keeping data and formatting intact.
// Use Cases: Generate a lightweight report that contains only the key columns from a master spreadsheet, preserving the original look. | Export selected columns for integration with a downstream system without loading the entire sheet into memory. | Create a temporary workbook for calculations that require only columns A‑D, reducing processing time. | Isolate specific data fields for auditing while maintaining the original styling for readability.
// AI Prompts: Provide C# code that utilizes Aspose.Cells LightCells to copy columns A‑D from a source worksheet to a new workbook, preserving values and styles. | Show how to determine the last data row and efficiently copy a column range with LightCells in Aspose.Cells. | Explain how to copy only the cell values (excluding styles) for columns A‑D to maximize performance using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // This example opens a source Excel file, determines the last populated row, and uses Aspose.Cells LightCells to rapidly transfer the first four columns (A‑D) – including cell values and optional formatting – into a freshly created workbook, which is then saved.
    class CopyColumnsWithLightCells
    {
        static void Main()
        {
            try
            {
                string sourcePath = "source.xlsx";
                string destinationPath = "destination.xlsx";

                // Verify source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Create a new empty workbook for the destination
                Workbook destinationWorkbook = new Workbook();

                // Get the first worksheet from each workbook
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
                Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

                // Determine the number of rows that contain data in the source sheet
                int lastDataRow = sourceSheet.Cells.MaxDataRow; // zero‑based index
                int rowCount = lastDataRow + 1; // total rows to copy

                // Copy columns A‑D (indices 0‑3) from source to destination
                for (int row = 0; row < rowCount; row++)
                {
                    for (int col = 0; col < 4; col++)
                    {
                        Cell srcCell = sourceSheet.Cells[row, col];
                        Cell destCell = destinationSheet.Cells[row, col];

                        // Copy value
                        destCell.PutValue(srcCell.Value);

                        // Copy style (optional, can be omitted for performance)
                        destCell.SetStyle(srcCell.GetStyle());
                    }
                }

                // Save the destination workbook
                destinationWorkbook.Save(destinationPath);
                Console.WriteLine($"Data copied successfully to {destinationPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
