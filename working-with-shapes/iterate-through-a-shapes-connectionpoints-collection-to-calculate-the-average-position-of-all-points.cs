// Title: Compute the centroid of a shape's connection points in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an Excel file with Aspose.Cells, gets the first shape on the first worksheet, extracts its ConnectionPoints via GetConnectionPoints, and returns the average X and Y coordinates. | Show a complete Aspose.Cells example that iterates over a shape's ConnectionPoints collection, sums the PointF positions, and prints the centroid of those points. | Demonstrate how to safely handle missing shapes or connection points while calculating the mean position of a shape's connection points and then save the workbook.
// Common Searches: aspocells calculate centroid of shape connection points c# | c# iterate shape connection points Aspose.Cells example | how to get average X Y of Excel shape connection points using Aspose.Cells | retrieve connection points positions from a shape in Aspose.Cells .NET
// Tags: aspocells shape connectionpoints iteration | c# calculate shape connection points centroid | aspocells getconnectionpoints average coordinates | excel workbook shape connection points analysis | dynamic connectionpoints handling aspocells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel workbook, accesses the first worksheet and its first shape, obtains the shape's connection points via GetConnectionPoints, iterates through them to sum X and Y values, computes the average (centroid) position, prints the result, and saves the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one shape
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the first worksheet.");
                return;
            }

            // Assume the shape we are interested in is the first shape on the sheet
            Shape shape = sheet.Shapes[0];

            // Retrieve connection points collection using dynamic to avoid compile‑time type issues
            dynamic connPoints = shape.GetConnectionPoints();

            // Ensure the shape has connection points
            if (connPoints != null && connPoints.Count > 0)
            {
                float sumX = 0f;
                float sumY = 0f;
                int count = connPoints.Count;

                // Iterate through all connection points
                foreach (dynamic cp in connPoints)
                {
                    // Position is a PointF (X, Y)
                    PointF pos = cp.Position;
                    sumX += pos.X;
                    sumY += pos.Y;
                }

                // Calculate average position
                float avgX = sumX / count;
                float avgY = sumY / count;

                // Output the result
                Console.WriteLine($"Average Position: X = {avgX}, Y = {avgY}");
            }
            else
            {
                Console.WriteLine("The selected shape has no connection points.");
            }

            // Save the workbook (optional if modifications were made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
