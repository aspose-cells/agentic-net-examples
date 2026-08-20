// Title: C# – Link a Rectangle Shape to a Worksheet Cell Using Aspose.Cells Shape.LinkedCell
// Description: Demonstrates how to create a workbook, add a rectangle shape, set its LinkedCell property to "$C$5", verify the link, and save the file as SmartArtLinkedCell.xlsx with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# shape linked cell | Shape.LinkedCell example | add rectangle shape Aspose.Cells | link shape to cell Excel .NET | Aspose.Cells Shape API | C# Excel shape anchoring
// Common Searches: Aspose.Cells how to link a shape to a cell | C# set LinkedCell property for rectangle shape | link shape to worksheet cell Aspose.Cells .NET | example of Shape.LinkedCell in C# | anchor shape to cell in Excel using Aspose
// Developer Intent: Attach a worksheet shape to a specific cell so the shape moves with the cell.
// Use Cases: Create a visual marker that stays aligned with a key data cell. | Build interactive dashboards where shapes navigate to linked cells. | Generate reports that automatically reposition shapes when rows or columns are inserted.
// AI Prompts: Generate C# code that adds a SmartArt shape and links it to cell D10 with Aspose.Cells. | Show how to change the value of a linked cell and confirm the shape follows the cell movement. | Provide a sample that links multiple shapes to different cells and saves the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartArtLinkExample
{
    // Demonstrates how to create a workbook, add a rectangle shape, set its LinkedCell property to "$C$5", verify the link, and save the file as SmartArtLinkedCell.xlsx with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet.
                // Parameters: shape type, upper left row, upper left column,
                //            row offset (pixels), column offset (pixels), height (pixels), width (pixels)
                Shape shape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // shape type
                    1,                       // upper left row (zero‑based)
                    1,                       // upper left column (zero‑based)
                    0,                       // row offset in pixels
                    0,                       // column offset in pixels
                    200,                     // height in pixels
                    200);                    // width in pixels

                // Link the shape to cell C5 (A1‑style address)
                shape.LinkedCell = "$C$5";

                // Verify the link
                Console.WriteLine("Shape linked to cell: " + shape.LinkedCell);

                // Save the workbook
                string outputPath = "SmartArtLinkedCell.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
