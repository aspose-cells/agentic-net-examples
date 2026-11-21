// Title: Add a rectangle shape to an Excel worksheet, display its type and position, and save the workbook with Aspose.Cells for .NET
// Common Searches: how to add a rectangle shape to an Excel worksheet using Aspose.Cells in C# | asp.net console application retrieve shape type and coordinates with Aspose.Cells | save an Aspose.Cells workbook to a specific folder after inserting a shape | ensure output directory exists before calling Workbook.Save in C# | example of using Shape.UpperLeftRow and Shape.UpperLeftColumn properties in Aspose.Cells | log shape information to console with Aspose.Cells .NET
// Tags: add rectangle shape Aspose.Cells | shape type property Aspose.Cells | retrieve shape coordinates Aspose.Cells | output directory creation before workbook save | Aspose.Cells workbook save example | console logging of shape details C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new Workbook, adds a rectangle shape at row 2 column 2, outputs the shape's type and upper‑left row/column to the console, ensures the target directory exists, and saves the workbook as output.xlsx while handling any exceptions.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: MsoDrawingType, upper left row, upper left column, top offset, left offset, height, width
                Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 200);

                // Output basic shape information using the correct properties
                Console.WriteLine($"Shape added. Type: {shape.Type}, UpperLeftRow: {shape.UpperLeftRow}, UpperLeftColumn: {shape.UpperLeftColumn}");

                // Save the workbook
                string outputPath = "output.xlsx";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
