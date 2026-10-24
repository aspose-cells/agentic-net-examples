// Title: Move the last shape to the back of the Z-order in an Excel worksheet and save the file with Aspose.Cells for .NET
// Common Searches: how to move the last shape to the back of the drawing order in an Excel sheet using Aspose.Cells C# | Aspose.Cells example for setting shape ZOrderPosition to the lowest value | C# code to reorder worksheet shapes and save the workbook with Aspose.Cells | change shape Z-order programmatically in an Excel file using Aspose.Cells .NET | move a shape behind other objects in Excel via Aspose.Cells API
// Tags: Aspose.Cells move shape Z-order | Aspose.Cells reorder worksheet shapes | C# set shape ZOrderPosition Aspose.Cells | Aspose.Cells save workbook after shape manipulation | Excel shape drawing order Aspose.Cells .NET

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an existing Excel workbook, checks for shapes on the first worksheet, retrieves the last shape, sets its ZOrderPosition to 0 to place it behind all other objects, ensures the output directory exists, and then saves the modified workbook, handling any errors that may occur.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Check for at least one shape on the sheet
            if (worksheet.Shapes.Count > 0)
            {
                // Get the last shape in the collection
                Shape lastShape = worksheet.Shapes[worksheet.Shapes.Count - 1];

                try
                {
                    // Move the shape to the back of the drawing order by setting Z-order to the lowest value
                    lastShape.ZOrderPosition = 0;
                }
                catch (Exception shapeEx)
                {
                    Console.WriteLine($"Failed to change shape Z-order: {shapeEx.Message}");
                }
            }
            else
            {
                Console.WriteLine("No shapes found on the worksheet.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
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
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
