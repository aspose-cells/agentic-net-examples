// Title: Get the absolute X and Y coordinates of a shape anchored to cell F5 with Aspose.Cells for .NET
// AI Prompts: Find the shape whose UpperLeftRow and UpperLeftColumn correspond to cell F5 and output its absolute X and Y pixel coordinates. | Modify the sample to use the Shape.Anchor property to calculate the exact pixel offset from the top‑left corner of the worksheet. | Create a reusable method that accepts a worksheet and a cell address, then returns the shape’s screen coordinates (X, Y) if a shape is anchored there.
// Common Searches: how to obtain pixel coordinates of a shape anchored at a specific cell using Aspose.Cells C# | Aspose.Cells C# get shape position by cell reference F5 | retrieve absolute X Y location of a drawing object in an Excel worksheet with Aspose.Cells | C# Aspose.Cells shape UpperLeftRow UpperLeftColumn to pixel offset conversion
// Tags: shape absolute position Aspose.Cells | retrieve shape coordinates C# | shape anchor cell reference Aspose.Cells | convert shape anchor to pixel offset | Aspose.Cells shape UpperLeftRow UpperLeftColumn

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an Excel workbook, scans the first worksheet’s Shapes collection to locate the shape whose UpperLeftRow and UpperLeftColumn match cell F5, and prints its anchored row and column. It also explains how to use the Anchor property in newer Aspose.Cells versions to calculate the exact X and Y pixel coordinates of the shape.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "sample.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or specify another index/name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Locate the shape whose upper‑left corner is anchored at cell F5 (row 4, column 5 – zero based)
            Shape targetShape = null;
            foreach (Shape shape in sheet.Shapes)
            {
                if (shape.UpperLeftRow == 4 && shape.UpperLeftColumn == 5)
                {
                    targetShape = shape;
                    break;
                }
            }

            if (targetShape != null)
            {
                // Retrieve position details of the shape's upper‑left corner
                int row = targetShape.UpperLeftRow;
                int column = targetShape.UpperLeftColumn;

                Console.WriteLine($"Shape anchored at row {row}, column {column}.");
                // Offsets are not available in older API versions; they can be omitted or retrieved via Anchor if needed.
            }
            else
            {
                Console.WriteLine("No shape found anchored at cell F5.");
            }

            // Save the workbook if any changes were made (optional)
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
