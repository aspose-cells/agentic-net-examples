// Title: Fit OleObject to a Cell Range – Set Width & Height with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to calculate the combined pixel width of a column range and the pixel height of a row range, add an OLE object at the range's top‑left cell, and set OleObject.Width, OleObject.Height, and Placement so the object exactly fills the selected cells. The workbook is then saved as OleObjectFitRange.xlsx.
// Keywords: Aspose.Cells C# OleObject size | fit OLE object to cell range | set OleObject Width Height Aspose | GetColumnWidthPixel Aspose.Cells | GetRowHeightPixel Aspose.Cells | OleObject Placement MoveAndSize | calculate pixel dimensions Aspose | Aspose.Cells add OLE object | C# spreadsheet OLE embedding
// Common Searches: Aspose.Cells set OleObject width and height | C# calculate total column width in pixels Aspose | fit OLE object across multiple cells Aspose.Cells | OleObject placement MoveAndSize example | how to make OLE object resize with cells Aspose
// Developer Intent: Resize an OLE object so it precisely covers a specified block of rows and columns by applying the summed pixel dimensions of those cells.
// Use Cases: Embed a chart that spans rows 2‑5 and columns B‑D and moves with the sheet. | Insert a PDF as an OLE object occupying a defined cell block for reporting templates. | Programmatically adjust OLE object dimensions after column width or row height changes.
// AI Prompts: Generate C# code using Aspose.Cells that adds an OLE object sized to a given start/end row and column range and updates its Width/Height when the worksheet layout changes. | Explain how to retrieve column widths and row heights in pixels with Aspose.Cells and apply them to OleObject dimensions. | Create a reusable method that accepts startRow, endRow, startColumn, endColumn and returns the pixel width and height needed for an OleObject.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to calculate the combined pixel width of a column range and the pixel height of a row range, add an OLE object at the range's top‑left cell, and set OleObject.Width, OleObject.Height, and Placement so the object exactly fills the selected cells. The workbook is then saved as OleObjectFitRange.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the target cell range where the OLE object should fit
            // Example: rows 2 to 5 (zero‑based) and columns 1 to 3
            int startRow = 2;
            int endRow = 5;
            int startColumn = 1;
            int endColumn = 3;

            // Calculate total width in pixels by summing column widths in the range
            int totalWidthPixels = 0;
            for (int col = startColumn; col <= endColumn; col++)
            {
                // Get column width in pixels directly
                double colWidthPixels = worksheet.Cells.GetColumnWidthPixel(col);
                totalWidthPixels += (int)Math.Ceiling(colWidthPixels);
            }

            // Calculate total height in pixels by summing row heights in the range
            int totalHeightPixels = 0;
            for (int row = startRow; row <= endRow; row++)
            {
                // Get row height in pixels directly
                double rowHeightPixels = worksheet.Cells.GetRowHeightPixel(row);
                totalHeightPixels += (int)Math.Ceiling(rowHeightPixels);
            }

            // Prepare image data for the OLE object (using an empty byte array for demo purposes)
            byte[] imageData = new byte[0];

            // Add the OLE object at the upper‑left cell of the range with the calculated size
            // Parameters: topRow, leftColumn, height (pixels), width (pixels), imageData
            int oleIndex = worksheet.OleObjects.Add(startRow, startColumn, totalHeightPixels, totalWidthPixels, imageData);
            OleObject oleObject = worksheet.OleObjects[oleIndex];

            // Ensure the OLE object's Width and Height exactly match the target cell range
            oleObject.Width = totalWidthPixels;
            oleObject.Height = totalHeightPixels;

            // Optional: set placement so the object moves and resizes with cells
            oleObject.Placement = PlacementType.MoveAndSize;

            // Save the workbook
            workbook.Save("OleObjectFitRange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
