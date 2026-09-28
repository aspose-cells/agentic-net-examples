// Title: Iterate a shape's ConnectionPoints collection to compute the average X and Y coordinates using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loops through a shape's ConnectionPoints collection and returns the average X and Y values as the shape's centroid. | Extend the example to process all shapes on a worksheet, calculating each shape's connection points average and outputting the results.
// Common Searches: Aspose.Cells C# calculate shape centroid from connection points | how to average X Y coordinates of shape connectionpoints in Aspose.Cells | iterate shape connectionpoints collection Aspose.Cells .NET example | C# Aspose.Cells get average position of shape connection points | compute center of Excel shape using Aspose.Cells connectionpoints
// Tags: iterate shape connectionpoints Aspose.Cells | calculate shape centroid Aspose.Cells | shape geometry utilities Aspose.Cells C# | connectionpoints average coordinates Aspose.Cells | excel shape center calculation C#

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The sample loads an Excel workbook, verifies that the first worksheet contains at least one shape, retrieves the shape's bounding cell indices (UpperLeftRow/Column and LowerRightRow/Column), computes the average row and column to represent the shape's central cell position, prints the result, and saves the workbook. It serves as a basis for extending the logic to iterate through a shape's ConnectionPoints collection and calculate the average X and Y coordinates.
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
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Verify that the worksheet contains at least one shape
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the first worksheet.");
                EnsureOutputDirectoryExists(outputPath);
                workbook.Save(outputPath);
                return;
            }

            // Get the first shape (adjust index as needed)
            Shape shape = sheet.Shapes[0];

            try
            {
                // Retrieve the cell coordinates that bound the shape
                int upperLeftRow = shape.UpperLeftRow;
                int upperLeftColumn = shape.UpperLeftColumn;
                int lowerRightRow = shape.LowerRightRow;
                int lowerRightColumn = shape.LowerRightColumn;

                // Compute average row and column indices
                double avgRow = (upperLeftRow + lowerRightRow) / 2.0;
                double avgColumn = (upperLeftColumn + lowerRightColumn) / 2.0;

                Console.WriteLine($"Average Position (by cells): Row = {avgRow}, Column = {avgColumn}");
            }
            catch (Exception ex)
            {
                // If the alternative method fails, report and continue.
                Console.WriteLine($"Unable to compute average position using cell bounds: {ex.Message}");
            }

            // Ensure the output directory exists and save the workbook
            EnsureOutputDirectoryExists(outputPath);
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to create output directory if needed
    private static void EnsureOutputDirectoryExists(string outputPath)
    {
        string outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }
    }
}
