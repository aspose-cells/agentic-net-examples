// Title: How to bind a rectangle shape's text to cell K3 for automatic updates using Aspose.Cells for .NET
// AI Prompts: Add a rectangle shape to the first worksheet and assign the formula "=K3" to its Text property for live cell synchronization. | Programmatically bind a shape's displayed text to a worksheet cell using Aspose.Cells C# and persist the changes by saving the workbook. | Create a shape, set its Text to a cell reference, and confirm that modifications to K3 automatically update the shape's content.
// Common Searches: Aspose.Cells C# link shape text to a specific cell value | How to make a shape display a cell's content and update automatically in .NET | Set shape text formula =K3 using Aspose.Cells library | Create rectangle shape with dynamic text from cell K3 in Excel via C# | Auto‑updating shape text in Aspose.Cells workbook
// Tags: shape.text formula binding Aspose.Cells | add rectangle shape C# Aspose.Cells | link shape to worksheet cell .NET | auto updating shape content Excel | Aspose.Cells dynamic shape text

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new workbook, adds a rectangle shape to the first worksheet, sets the shape's Text property to the formula "=K3" so it mirrors the value of cell K3 and updates automatically, ensures the output directory exists, and saves the file as LinkedShape.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet and obtain the Shape object directly
                Shape shape = sheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // shape type
                    1,   // upper left row
                    1,   // upper left column
                    0,   // top offset (pixels)
                    0,   // left offset (pixels)
                    100, // height (pixels)
                    50   // width (pixels)
                );

                // Link the shape's text to cell K3 (the shape will display the value of K3 and update automatically)
                shape.Text = "=K3";

                // Define output path
                string outputPath = "LinkedShape.xlsx";

                // Ensure the directory exists (handle case when outputPath has no directory part)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
