// Title: Retrieve a shape's absolute position in points, convert to centimeters, and write coordinates to Excel cells using Aspose.Cells for .NET
// Common Searches: how to get the X and Y position of a shape in an Excel worksheet with Aspose.Cells C# | convert shape.Left and shape.Top from points to centimeters in Aspose.Cells | store shape coordinates in cells A1 and B1 using Aspose.Cells for .NET | Aspose.Cells C# example for shape position conversion to metric units | calculate shape location in centimeters from points in an Excel file with Aspose.Cells
// Tags: convert shape points to centimeters Aspose.Cells | write shape location to worksheet cells C# | retrieve shape left top values Aspose.Cells API | pixel to centimeter conversion using DPI reference | C# metric conversion of Excel shape location

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads (or creates) an Excel workbook, accesses the first worksheet, obtains the first shape, converts its Left and Top values from points to centimeters using a 96 dpi reference, writes the rounded X and Y values to cells A1 and B1, and saves the workbook.
class ShapePositionToCentimeters
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load workbook if it exists; otherwise create a new one.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one shape on the worksheet.
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the worksheet.");
                workbook.Save(outputPath);
                return;
            }

            // Get the first shape (adjust index as needed).
            Shape shape = sheet.Shapes[0];

            // Convert shape position from points to centimeters.
            const double dpi = 96.0;
            double pointsToPixels = dpi / 72.0;   // 1 point = 1/72 inch.
            double pixelToCm = 2.54 / dpi;        // 1 pixel = 2.54 / 96 cm.

            // Shape.Left and Shape.Top are measured in points.
            double xPixels = shape.Left * pointsToPixels;
            double yPixels = shape.Top * pointsToPixels;

            double xCentimeters = xPixels * pixelToCm;
            double yCentimeters = yPixels * pixelToCm;

            // Write the converted values to cells A1 (X) and B1 (Y).
            Cells cells = sheet.Cells;
            cells["A1"].PutValue(Math.Round(xCentimeters, 4));
            cells["B1"].PutValue(Math.Round(yCentimeters, 4));

            // Save the workbook with the new data.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
