// Title: Read a shape's width, height and top‑left cell coordinates and write them to Excel cells using Aspose.Cells for .NET
// Common Searches: how to read shape width and height in an Excel file with Aspose.Cells C# | Aspose.Cells C# write shape dimensions to specific worksheet cells | retrieve shape upper left row and column indices using Aspose.Cells .NET | calculate bounding box of a shape in a workbook with Aspose.Cells | save shape size information to cells A1 and B1 using Aspose.Cells
// Tags: Aspose.Cells shape dimension extraction | Aspose.Cells write shape size to cells | Aspose.Cells bounding box calculation | C# read shape position indices | C# write numeric values to worksheet cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Required for Shape class

// The example loads an existing workbook, accesses the first worksheet, reads the first shape's Width, Height, UpperLeftRow, and UpperLeftColumn properties, writes the width and height to cells A1 and B1 and the position to A2 and B2, and saves the modified workbook as a new file.
class ShapeBoundingBoxExample
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one shape
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the worksheet.");
                return;
            }

            // Retrieve the first shape (or select by name if required)
            Shape shape = sheet.Shapes[0];

            // Use the shape's Width and Height properties as the bounding box dimensions
            double width = shape.Width;
            double height = shape.Height;

            // Obtain the shape's top‑left position (row/column)
            int upperLeftRow = shape.UpperLeftRow;
            int upperLeftColumn = shape.UpperLeftColumn;

            // Write dimensions to cells
            sheet.Cells["A1"].PutValue(width);
            sheet.Cells["B1"].PutValue(height);

            // Write position information for reference
            sheet.Cells["A2"].PutValue(upperLeftRow);
            sheet.Cells["B2"].PutValue(upperLeftColumn);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
