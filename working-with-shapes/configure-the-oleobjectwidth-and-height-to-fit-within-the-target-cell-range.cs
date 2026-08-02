// Title: Fit an OleObject to a Cell Range with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to add an OleObject to a workbook and anchor it to a specific cell block (e.g., B2:D5) by setting UpperLeftRow/Column and LowerRightRow/Column, ensuring the object's width and height match the target range.
// Keywords: Aspose.Cells OleObject resize | C# fit OleObject to cells | anchor OleObject cell range | set OleObject dimensions Aspose | OleObject UpperLeftRow LowerRightRow
// Common Searches: Aspose.Cells resize OleObject to cell range | C# set OleObject width height Aspose | fit embedded OLE object to B2:D5 | how to anchor OleObject to specific cells in .NET | adjust OleObject dimensions programmatically
// Developer Intent: Resize an OleObject so it exactly covers a defined cell range.
// Use Cases: Generate reports where a placeholder image automatically fills a table area. | Align an embedded chart with a data block so it scales with surrounding cells. | Embed a Word document that matches the size of a designated cell region in a spreadsheet.
// AI Prompts: Write C# code that calculates OleObject.Width and Height from column widths and row heights to fit a given range in Aspose.Cells. | Show how to modify an existing OleObject after loading a workbook so it aligns with cells C3:E6. | Explain how to retrieve column width and row height values in Aspose.Cells to set an OleObject's size accurately.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add an OleObject to a workbook and anchor it to a specific cell block (e.g., B2:D5) by setting UpperLeftRow/Column and LowerRightRow/Column, ensuring the object's width and height match the target range.
class OleObjectFitToCellRange
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the target cell range (B2:D5)
            int startRow = 1;      // B2 -> row index 1 (zero‑based)
            int endRow = 4;        // D5 -> row index 4
            int startColumn = 1;   // Column B -> index 1
            int endColumn = 3;     // Column D -> index 3

            // Add a placeholder OleObject (size will be adjusted later)
            // Use a dummy 1x1 pixel transparent PNG to satisfy the API
            byte[] dummyImage = new byte[]
            {
                137,80,78,71,13,10,26,10,0,0,0,13,73,72,68,82,
                0,0,0,1,0,0,0,1,8,6,0,0,0,31,21,196,
                137,0,0,0,12,73,68,65,84,8,153,99,0,1,0,0,
                5,0,1,0,0,0,0,73,69,78,68,174,66,96,130
            };
            int oleIndex = sheet.OleObjects.Add(startRow, startColumn, 0, 0, dummyImage);
            OleObject ole = sheet.OleObjects[oleIndex];

            // Anchor the OleObject to the exact cell range
            ole.UpperLeftRow = startRow;
            ole.UpperLeftColumn = startColumn;
            ole.LowerRightRow = endRow;
            ole.LowerRightColumn = endColumn;

            // Determine output path and ensure its directory exists
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleObjectFitToRange.xlsx");
            string outputDir = Path.GetDirectoryName(outputPath);
            if (string.IsNullOrEmpty(outputDir))
            {
                outputDir = Directory.GetCurrentDirectory();
            }
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
