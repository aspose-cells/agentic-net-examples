// Title: Read a shape's absolute position in points, convert to millimeters, and store the values in hidden Excel cells using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get shape top and left coordinates in points | convert shape position from points to millimeters with Aspose.Cells | write shape coordinates to a hidden column in Excel using Aspose.Cells | hide worksheet column after storing shape data programmatically Aspose.Cells | read absolute position of a shape in an Excel workbook with Aspose.Cells
// Tags: shape absolute position points to millimeters Aspose.Cells | write values to hidden Excel column Aspose.Cells | convert point units to millimeter .NET | access shape Top Left properties Aspose.Cells | hide worksheet column programmatically Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an existing workbook, retrieves the first shape's Top and Left coordinates in points, converts those measurements to millimeters (1 point = 25.4/72 mm), writes the converted values to hidden cells (Z1 and Z2), hides the column containing those cells, and saves the updated workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one shape on the worksheet
            if (worksheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the first worksheet.");
                return;
            }

            // Get the first shape (adjust index or name as needed)
            Shape shape = worksheet.Shapes[0];

            // Read the shape's absolute position in points
            double topPoints = shape.Top;   // distance from the top edge of the worksheet
            double leftPoints = shape.Left; // distance from the left edge of the worksheet

            // Convert points to millimeters (1 point = 25.4 / 72 mm)
            const double pointToMillimeter = 25.4 / 72.0;
            double topMillimeters = topPoints * pointToMillimeter;
            double leftMillimeters = leftPoints * pointToMillimeter;

            // Write the converted values to hidden cells (e.g., Z1 for Top, Z2 for Left)
            Cell topCell = worksheet.Cells["Z1"];
            topCell.PutValue(topMillimeters);
            Cell leftCell = worksheet.Cells["Z2"];
            leftCell.PutValue(leftMillimeters);

            // Hide the column containing the hidden cells
            int hiddenColumnIndex = topCell.Column;
            worksheet.Cells.HideColumn(hiddenColumnIndex);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
