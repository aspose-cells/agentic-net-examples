// Title: C# Aspose.Cells example: Read shape connection points and log those on the shape's right edge
// Common Searches: Aspose.Cells C# get connection points of a shape in an Excel worksheet | How to identify right‑edge connection points of Excel shapes using Aspose.Cells | Iterate over worksheet shapes and log connection points on the right side with Aspose.Cells for .NET | C# code to compare shape LowerRightColumn with connection point column in Aspose.Cells | Read and output shape connection point coordinates from an Excel file using Aspose.Cells
// Tags: Aspose.Cells shape connection points C# | right edge shape connection detection Aspose.Cells | iterate worksheet shapes Aspose.Cells .NET | shape lower right column comparison Aspose.Cells | log shape connection coordinates C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel workbook (or creates a new one), ensures a worksheet exists, iterates through all shapes on the first worksheet, retrieves each shape's connection points via GetConnectionPoints(), checks whether a point's column matches the shape's LowerRightColumn (the right edge), writes matching points to the console, and finally saves the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Load existing workbook or create a new one if the file is missing
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
                workbook.Worksheets.Add();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all shapes on the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                try
                {
                    // Column index of the shape's right edge
                    int rightEdgeColumn = shape.LowerRightColumn;

                    // Get the collection of connection points for the shape (use dynamic to avoid compile‑time type issues)
                    dynamic cpCollection = shape.GetConnectionPoints();

                    // Examine each connection point of the shape
                    foreach (dynamic cp in cpCollection)
                    {
                        // Check if the connection point lies on the right edge
                        if ((int)cp.Column == rightEdgeColumn)
                        {
                            Console.WriteLine(
                                $"Shape '{shape.Name}' has a right‑edge connection point at Row={cp.Row}, Column={cp.Column}");
                        }
                    }
                }
                catch (Exception shapeEx)
                {
                    Console.WriteLine($"Error processing shape '{shape.Name}': {shapeEx.Message}");
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook (optional if modifications were made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
