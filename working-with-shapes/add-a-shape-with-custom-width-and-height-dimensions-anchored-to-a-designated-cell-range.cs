// Title: Add a rectangle shape with custom width and height anchored to cells B2:D5 using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a new workbook, adds a rectangle shape anchored to the range B2:D5, and sets its width to 200 points and height to 100 points with Aspose.Cells. | Show how to use Aspose.Cells' Shapes.AddShape method to place a rectangle in a specific cell range and then modify its Size properties programmatically. | Demonstrate adjusting the Width and Height of a shape after anchoring it to a worksheet range in a .NET application using Aspose.Cells.
// Common Searches: Aspose.Cells C# add rectangle shape to specific cell range and set size | how to set custom width and height for a shape anchored to cells in Aspose.Cells .NET | example of using Shapes.AddShape with B2:D5 range in Aspose.Cells | change shape dimensions after adding to worksheet Aspose.Cells C# | anchor shape to cell range and resize in Aspose.Cells for .NET
// Tags: add rectangle shape Aspose.Cells | anchor shape to cell range | set shape width height points | Shapes.AddShape usage .NET | custom shape dimensions Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing; // For Color if needed

// The example creates a new workbook, selects the first worksheet, defines the cell range B2:D5, adds a rectangle shape anchored to that range, sets the shape's width to 200 points and height to 100 points, and saves the workbook as ShapeWithCustomSize.xlsx, with exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the cell range to anchor the shape (B2:D5)
            int startRow = 1;      // B2 row (zero‑based)
            int startColumn = 1;   // B2 column (zero‑based)
            int endRow = 4;        // D5 row (zero‑based)
            int endColumn = 3;     // D5 column (zero‑based)

            // Add a rectangle shape anchored to the specified range
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // Correct enum usage
                startRow, startColumn,    // upper‑left cell
                0, 0,                     // offset within the upper‑left cell
                endRow, endColumn);       // lower‑right cell

            // Set custom dimensions (points)
            shape.Width = 200;   // width
            shape.Height = 100;  // height

            // Optional visual styling (uncomment if supported)
            // shape.LineWeight = 1.0;
            // shape.FillColor = Color.LightBlue;

            // Save the workbook
            workbook.Save("ShapeWithCustomSize.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
