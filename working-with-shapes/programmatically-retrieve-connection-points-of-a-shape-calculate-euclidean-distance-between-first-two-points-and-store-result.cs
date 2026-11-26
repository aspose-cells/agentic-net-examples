// Title: Retrieve shape connection points, compute Euclidean distance between the first two points, and write the result to an Excel cell using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get shape connection points and calculate distance | How to compute Euclidean distance between two shape points in an Excel worksheet with Aspose.Cells | Write a calculated numeric value to cell A1 using Aspose.Cells .NET | C# example for extracting connection points from the first shape in a workbook | Save a workbook after updating cell values with Aspose.Cells
// Tags: shape GetConnectionPoints Aspose.Cells .NET | calculate Euclidean distance C# Excel | write value to worksheet cell Aspose.Cells | access first shape on worksheet Aspose.Cells | save modified workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an existing Excel file, accesses the first worksheet, retrieves the first shape's connection points via GetConnectionPoints, calculates the Euclidean distance between the first two points, writes the distance (or an error message) into cell A1, and saves the workbook as a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one shape
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the worksheet.");
                return;
            }

            Shape shape = sheet.Shapes[0];

            // Get all connection points of the shape (returned as float[][]: each inner array is {X, Y})
            float[][] points = shape.GetConnectionPoints();

            // Ensure there are at least two points with valid coordinates
            if (points != null && points.Length >= 2 &&
                points[0] != null && points[0].Length >= 2 &&
                points[1] != null && points[1].Length >= 2)
            {
                // Calculate Euclidean distance between the first two points
                double dx = points[1][0] - points[0][0]; // X1 - X0
                double dy = points[1][1] - points[0][1]; // Y1 - Y0
                double distance = Math.Sqrt(dx * dx + dy * dy);

                // Store the result in cell A1
                sheet.Cells["A1"].PutValue(distance);
            }
            else
            {
                // If insufficient points, write a message
                sheet.Cells["A1"].PutValue("Not enough connection points");
            }

            // Save the workbook with the calculated result
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
