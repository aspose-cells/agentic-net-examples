// Title: Read a shape's current text alignment and center it using Aspose.Cells for .NET (C#)
// Common Searches: C# code to get the current horizontal and vertical text alignment of an Excel shape with Aspose.Cells | Aspose.Cells example for reading shape text alignment before modifying it | How to center shape text after retrieving its existing alignment in a .NET workbook | Retrieve and update shape text alignment in an existing XLSX file using Aspose.Cells C# | Check shape text alignment values in Aspose.Cells before applying changes
// Tags: Aspose.Cells read shape text alignment C# | update shape text horizontal vertical alignment Aspose.Cells | center shape text alignment Aspose.Cells .NET | retrieve shape text properties Excel workbook C# | modify shape text alignment in existing XLSX using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an existing XLSX workbook, verifies that the first worksheet contains a shape, reads the shape's current horizontal and vertical text alignment, prints the values, then sets both alignments to Center, saves the workbook to a new file, and logs the operation steps.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Verify that the worksheet contains at least one shape
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the first worksheet.");
                return;
            }

            // Retrieve the first shape on the worksheet
            Shape shape = sheet.Shapes[0];

            try
            {
                // Read current horizontal and vertical alignment
                TextAlignmentType currentHAlign = shape.TextHorizontalAlignment;
                TextAlignmentType currentVAlign = shape.TextVerticalAlignment;

                Console.WriteLine($"Current horizontal alignment: {currentHAlign}");
                Console.WriteLine($"Current vertical alignment: {currentVAlign}");

                // Apply new alignment values (example: center both horizontally and vertically)
                shape.TextHorizontalAlignment = TextAlignmentType.Center;
                shape.TextVerticalAlignment = TextAlignmentType.Center;

                Console.WriteLine("Alignment updated to Center (both horizontal and vertical).");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while modifying shape alignment: {ex.Message}");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
