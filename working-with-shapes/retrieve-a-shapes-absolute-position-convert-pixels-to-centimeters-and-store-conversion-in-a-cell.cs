// Title: Get a shape’s absolute position in pixels, convert to centimeters, and write the values to Excel cells using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# retrieve shape upper‑left column and row as pixel offsets | convert shape pixel coordinates to centimeters with Aspose.Cells drawing API | write converted shape X and Y positions to specific cells in an Excel workbook using Aspose.Cells | calculate centimeters from shape pixel offsets based on 96 dpi in .NET | store shape position values as cm in Excel cells with Aspose.Cells
// Tags: Aspose.Cells shape absolute position extraction | pixel offset to cm conversion using DPI | populate Excel cells with shape coordinates C# | Aspose.Cells drawing API coordinate handling | Excel workbook shape location conversion .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing workbook, accesses the first worksheet, reads the first shape's UpperLeftColumn and UpperLeftRow values (treated as pixel offsets), converts those offsets to centimeters using a 96 dpi to 2.54 cm conversion, writes the X and Y results into cells B2 and B3 with labels in A2 and A3, and saves the modified workbook to a new file.
class ShapePositionToCentimeters
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one shape
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the worksheet.");
                return;
            }

            // Get the first shape (adjust index if needed)
            Shape shape = sheet.Shapes[0];

            // Retrieve the shape's position (row/column indices are used as a fallback)
            int offsetXInPixels = 0;
            int offsetYInPixels = 0;

            try
            {
                // UpperLeftColumn and UpperLeftRow are always available
                offsetXInPixels = shape.UpperLeftColumn;
                offsetYInPixels = shape.UpperLeftRow;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Unable to read shape position ({ex.Message}). Offsets set to 0.");
            }

            // Convert the (approximate) pixel values to centimeters.
            // 1 inch = 2.54 cm, 1 inch = 96 pixels (standard screen DPI)
            const double pixelsPerInch = 96.0;
            const double cmPerInch = 2.54;

            double offsetXInCm = offsetXInPixels * cmPerInch / pixelsPerInch;
            double offsetYInCm = offsetYInPixels * cmPerInch / pixelsPerInch;

            // Store the conversion results in cells
            sheet.Cells["B2"].PutValue(offsetXInCm);
            sheet.Cells["B3"].PutValue(offsetYInCm);

            // Optional labels for clarity
            sheet.Cells["A2"].PutValue("Shape X (cm):");
            sheet.Cells["A3"].PutValue("Shape Y (cm):");

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the new data
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
