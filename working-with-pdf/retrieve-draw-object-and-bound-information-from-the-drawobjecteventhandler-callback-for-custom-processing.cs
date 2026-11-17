// Title: C# example for obtaining shape bounds and type from a worksheet using Aspose.Cells
// Common Searches: Aspose.Cells C# get rectangle shape dimensions from a worksheet | how to read shape upper left row and column in Aspose.Cells | retrieve shape width and height programmatically with Aspose.Cells .NET | access worksheet shape properties before saving workbook Aspose.Cells | C# Aspose.Cells example to list shape type and bounds
// Tags: Aspose.Cells retrieve worksheet shape bounds | C# get shape dimensions Aspose.Cells | Aspose.Cells shape property extraction | worksheet shape information Aspose.Cells .NET | Aspose.Cells add rectangle shape example

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an existing Excel file, adds a rectangle shape to the first worksheet, outputs the shape's type, upper‑left row and column, width, and height, and then saves the workbook, demonstrating how to obtain shape bounds and properties with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Add a rectangle shape to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            // Parameters: drawing type, upper left row, upper left column, row offset, column offset, width, height
            Shape rect = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 50);

            // Output basic information about the created shape
            Console.WriteLine($"Shape Type: {rect.GetType().Name}");
            Console.WriteLine($"UpperLeftRow: {rect.UpperLeftRow}, UpperLeftColumn: {rect.UpperLeftColumn}");
            Console.WriteLine($"Width: {rect.Width}, Height: {rect.Height}");

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
