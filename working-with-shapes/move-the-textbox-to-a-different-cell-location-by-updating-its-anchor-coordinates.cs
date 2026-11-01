// Title: Reposition an Excel TextBox to a new cell range by updating its anchor coordinates with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# move textbox to a specific cell range | How to set UpperLeftRow and UpperLeftColumn for a shape in Aspose.Cells | Change position of an Excel textbox programmatically using Aspose.Cells | Set shape anchor to B2:C5 with Aspose.Cells .NET | Update shape coordinates in an existing workbook using Aspose.Cells
// Tags: Aspose.Cells shape anchor positioning | C# Excel textbox relocation | Aspose.Cells CellArea shape bounds | Set UpperLeftRow UpperLeftColumn Aspose.Cells | Programmatic shape movement in workbook | Aspose.Cells shape coordinate update

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a workbook, retrieves the first shape (assumed to be a TextBox), defines a new CellArea covering cells B2:C5, assigns the shape's UpperLeftRow/UpperLeftColumn and LowerRightRow/LowerRightColumn to those coordinates, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Get the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Retrieve the first shape (covers TextBox, etc.)
            Shape shape = worksheet.Shapes.Count > 0 ? worksheet.Shapes[0] : null;

            if (shape != null)
            {
                // Define the new anchor area for the shape (e.g., cells B2:C5)
                var newAnchor = new CellArea
                {
                    StartRow = 1,    // Row 2 (zero‑based)
                    StartColumn = 1, // Column B (zero‑based)
                    EndRow = 4,      // Row 5
                    EndColumn = 2    // Column C
                };

                try
                {
                    // Set the shape's anchor using row/column properties
                    shape.UpperLeftRow = newAnchor.StartRow;
                    shape.UpperLeftColumn = newAnchor.StartColumn;
                    shape.LowerRightRow = newAnchor.EndRow;
                    shape.LowerRightColumn = newAnchor.EndColumn;
                    // Offsets are not required; they default to zero.
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to set shape anchor: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("No shape found in the worksheet.");
            }

            // Ensure the output directory exists
            var outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
