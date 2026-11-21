// Title: List all shapes on the first worksheet and output their UpperLeftRow and UpperLeftColumn positions using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get shape position UpperLeftRow UpperLeftColumn in Excel worksheet | How to iterate through worksheet shapes and read their coordinates with Aspose.Cells .NET | C# code to log each shape's name and top‑left cell address using Aspose.Cells | Retrieve shape coordinates from an existing .xlsx file using Aspose.Cells for .NET | Aspose.Cells enumerate shapes and print location information in console
// Tags: Aspose.Cells shape position extraction | enumerate worksheet shapes .NET | read shape UpperLeftRow UpperLeftColumn | log shape coordinates Excel | C# Aspose.Cells shape iteration

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing workbook, accesses the first worksheet, iterates through all shapes in the sheet, and writes each shape's name along with its UpperLeftRow and UpperLeftColumn values to the console before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the collection of shapes on the worksheet
            ShapeCollection shapes = sheet.Shapes;

            // Iterate through each shape
            foreach (Shape shape in shapes)
            {
                try
                {
                    // Retrieve basic position information for the shape
                    int upperRow = shape.UpperLeftRow;
                    int upperColumn = shape.UpperLeftColumn;

                    // Log shape information
                    Console.WriteLine($"Shape: {shape.Name}, UpperLeftRow: {upperRow}, UpperLeftColumn: {upperColumn}");
                }
                catch (Exception exShape)
                {
                    Console.WriteLine($"Error processing shape \"{shape.Name}\": {exShape.Message}");
                }
            }

            // Save the workbook after processing
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
