// Title: Retrieve a worksheet shape's absolute coordinates, convert from points to millimeters, and store them in hidden cells using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook, obtains the Left and Top properties of the first shape on a worksheet, converts those point values to millimeters, writes the converted numbers into hidden cells, and saves the workbook with Aspose.Cells. | Demonstrate how to hide a specific column after inserting shape position data into cells in an Excel file using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# convert shape Left and Top from points to millimeters | how to read absolute position of a shape in an Excel worksheet with Aspose.Cells | store shape coordinates in hidden Excel cells using Aspose.Cells .NET | hide column after writing data with Aspose.Cells C# | convert Excel shape coordinates to metric units in C#
// Tags: shape position metric conversion Aspose.Cells | write hidden cell values Aspose.Cells C# | hide column programmatically Aspose.Cells | retrieve shape absolute coordinates Aspose.Cells | excel shape coordinate conversion to metric units

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an existing workbook, reads the Left and Top coordinates (in points) of the first shape on the first worksheet, converts those values to millimeters, writes the results to hidden cells A1 and A2, hides column A, and saves the modified workbook.
class ShapePositionToMillimeters
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one shape on the sheet
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the worksheet.");
                return;
            }

            // Get the first shape (replace index if you need a specific shape)
            Shape shape = sheet.Shapes[0];

            // Shape.Left and Shape.Top return coordinates in points (1 point = 1/72 inch)
            double xPoints = shape.Left;
            double yPoints = shape.Top;

            // Convert points to millimeters: 1 point = 25.4 / 72 mm
            const double pointsToMillimeters = 25.4 / 72.0;
            double xMillimeters = xPoints * pointsToMillimeters;
            double yMillimeters = yPoints * pointsToMillimeters;

            // Write the converted values to hidden cells (A1 and A2)
            Cell cellX = sheet.Cells["A1"];
            Cell cellY = sheet.Cells["A2"];
            cellX.PutValue(xMillimeters);
            cellY.PutValue(yMillimeters);

            // Hide column A (index 0) so the cells are not visible to the user
            sheet.Cells.HideColumn(0);

            // Save the workbook with the new hidden data
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
