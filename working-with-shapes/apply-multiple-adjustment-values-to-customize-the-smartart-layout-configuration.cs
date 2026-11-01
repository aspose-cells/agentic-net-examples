// Title: Create a rectangle shape as a SmartArt placeholder, set its row/column position with pixel offsets and size, then save the workbook as XLSX using Aspose.Cells for .NET (C#)
// Common Searches: how to add a rectangle shape to an Excel worksheet with Aspose.Cells C# | Aspose.Cells set shape row and column offsets in pixels | define shape height and width in pixels using Aspose.Cells API | save an Excel workbook containing a shape as .xlsx with Aspose.Cells | use Aspose.Cells to create a SmartArt placeholder shape in C#
// Tags: rectangle shape addition Aspose.Cells | shape position offsets Aspose.Cells | shape pixel dimensions Aspose.Cells | shape text assignment Aspose.Cells | workbook save xlsx Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, accesses the first worksheet, adds a rectangle shape as a SmartArt placeholder at a specific row and column with pixel offsets, sets its height and width in pixels, assigns the text "Start" to the shape, and saves the workbook as an XLSX file.
class SmartArtAdjustmentExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Define the shape position (row/column) and size (in pixels).
            int upperLeftRow = 2;
            int upperLeftColumn = 2;
            int upperLeftRowOffset = 0;   // offset within the cell (pixels)
            int upperLeftColumnOffset = 0;
            int shapeHeight = 200;        // height in pixels
            int shapeWidth = 300;         // width in pixels

            // Add a rectangle shape (as a placeholder for SmartArt – Aspose.Cells does not expose SmartArt directly).
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                upperLeftRow,
                upperLeftColumn,
                upperLeftRowOffset,
                upperLeftColumnOffset,
                shapeHeight,
                shapeWidth);

            // Optionally set text inside the shape.
            shape.Text = "Start";

            // Save the workbook to a file.
            string outputPath = "SmartArtAdjustmentResult.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
