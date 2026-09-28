// Title: Get a shape's absolute position in points, convert to centimeters, and write the coordinates to cells with Aspose.Cells for .NET
// AI Prompts: Using Aspose.Cells for .NET, load a workbook, find a shape by its name, read its Top and Left properties (points), convert those values to centimeters, and place the X and Y results into cells B1 and B2. | Write C# code that extracts a shape's absolute coordinates, transforms the point measurements to cm using the 72‑point‑per‑inch and 96‑pixel‑per‑inch ratios, and saves the converted values back into the worksheet.
// Common Searches: Aspose.Cells C# get shape top left position in centimeters | convert shape coordinates from points to cm in Excel using Aspose.Cells | write shape absolute location to specific cells with Aspose.Cells .NET | how to calculate shape position in cm from points in Aspose.Cells | Aspose.Cells retrieve shape position and store in worksheet cells
// Tags: Aspose.Cells shape top left points to centimeters conversion | C# extract absolute shape coordinates Aspose.Cells | store shape X Y values in worksheet cells Aspose.Cells | convert Excel shape measurements points to cm Aspose.Cells | write shape location into cells B1 B2 using Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an existing workbook, locates a shape named "MyShape", reads its Top and Left coordinates expressed in points, converts those measurements to centimeters, writes the X (left) value to cell B1 and the Y (top) value to cell B2, and saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the shape by name (adjust the name as needed)
            Shape shape = sheet.Shapes["MyShape"];
            if (shape == null)
            {
                Console.WriteLine("Error: Shape \"MyShape\" not found in the worksheet.");
                return;
            }

            // Use Shape.Top and Shape.Left (points) for absolute position
            double topPoints = shape.Top;   // distance from top of worksheet in points
            double leftPoints = shape.Left; // distance from left of worksheet in points

            // Convert points to pixels (1 point = 1/72 inch, 1 inch = 96 pixels)
            double topPixels = topPoints * 96.0 / 72.0;
            double leftPixels = leftPoints * 96.0 / 72.0;

            // Convert pixels to centimeters (1 inch = 96 pixels, 1 inch = 2.54 cm)
            double xCm = leftPixels * 2.54 / 96.0;
            double yCm = topPixels * 2.54 / 96.0;

            // Store the converted values in cells (e.g., B1 for X, B2 for Y)
            sheet.Cells["B1"].PutValue(xCm);
            sheet.Cells["B2"].PutValue(yCm);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the new data
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
