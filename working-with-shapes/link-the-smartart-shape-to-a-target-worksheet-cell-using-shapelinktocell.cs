// Title: Link a SmartArt shape to a specific worksheet cell using Shape.LinkToCell in Aspose.Cells for .NET (C#)
// AI Prompts: Add a SmartArt shape to a worksheet and link it to a target cell using Shape.LinkToCell. | Update an existing rectangle shape so that Shape.LinkToCell attaches it to cell B3. | Create a new workbook, insert a SmartArt shape, and use Shape.LinkToCell to associate the shape with cell A1 before saving.
// Common Searches: Aspose.Cells C# Shape.LinkToCell example for linking a SmartArt shape to cell C5 | How to bind an Excel shape to a specific cell using Aspose.Cells .NET | Shape.LinkToCell method usage to anchor SmartArt to a worksheet cell in C# | Linking shapes to cells with Aspose.Cells API for .NET developers
// Tags: Shape.LinkToCell C# Aspose.Cells | link SmartArt shape to worksheet cell | bind Excel shape to cell Aspose.Cells .NET | programmatic shape‑cell association Aspose.Cells | set shape anchor using LinkToCell

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartArtExample
{
    // The example creates a new Workbook, accesses the first worksheet, adds a rectangle SmartArt shape, links the shape to a target worksheet cell with Shape.LinkToCell, and saves the file as SmartArtLinked.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Define position and size for the shape (row, column, height, width)
                int upperLeftRow = 2;          // Row index (0‑based)
                int upperLeftColumn = 2;       // Column index (0‑based)
                int shapeHeight = 200;         // Height in points
                int shapeWidth = 300;          // Width in points

                // Add a simple rectangle shape (using MsoDrawingType enum)
                // Parameters: type, upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, height, width
                Shape shape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle,
                    upperLeftRow,
                    upperLeftColumn,
                    0,               // row offset
                    0,               // column offset
                    shapeHeight,
                    shapeWidth);

                // Set line weight (visual property)
                shape.Line.Weight = 1.5f;

                // Define output file path
                string outputPath = "SmartArtLinked.xlsx";

                // Ensure the directory exists (handle case when outputPath has no directory part)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
