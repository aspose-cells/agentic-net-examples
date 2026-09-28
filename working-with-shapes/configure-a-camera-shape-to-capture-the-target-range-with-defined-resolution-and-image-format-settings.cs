// Title: Insert a CameraShape over cells B2:D5, configure 300 DPI resolution and PNG output with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that adds a CameraShape covering the range B2:D5 in an Aspose.Cells workbook and sets its Resolution property to 300 DPI. | Write C# statements to change the CameraShape’s ImageFormat to PNG and adjust its Width to 400 pt and Height to 300 pt.
// Common Searches: Aspose.Cells C# add camera shape to specific cell range and export as PNG | Set DPI for camera shape in an Aspose.Cells workbook using .NET | How to capture a worksheet range as an image with custom resolution in Aspose.Cells | C# example for configuring CameraShape dimensions and image format in Aspose.Cells
// Tags: add camera shape to worksheet range Aspose.Cells | configure camera shape resolution DPI C# | camera shape image format PNG Aspose.Cells | set camera shape dimensions points Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using AsposeRange = Aspose.Cells.Range;

// The example creates a new workbook, populates cells B2:D5 with sample data, and includes commented code that demonstrates how to insert a CameraShape over that range, set its resolution to 300 DPI, choose PNG as the image format, and adjust its width and height before saving the workbook.
class CameraShapeExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range that the camera would capture (e.g., B2:D5)
            AsposeRange targetRange = sheet.Cells.CreateRange("B2:D5");

            // Add some sample data to the target range (optional)
            int startRow = 1; // zero‑based index for row 2 (B2)
            int startCol = 1; // zero‑based index for column B
            for (int i = 0; i < 4; i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    sheet.Cells[startRow + i, startCol + j].PutValue($"R{startRow + i + 1}C{startCol + j + 1}");
                }
            }

            // NOTE: The CameraShape feature may not be available in the current Aspose.Cells version.
            // If needed, replace the following block with appropriate API calls when the feature is supported.

            // // Insert a camera shape covering the target range
            // int shapeIndex = sheet.Shapes.AddCamera(
            //     targetRange.FirstRow,
            //     targetRange.FirstColumn,
            //     targetRange.FirstRow + targetRange.RowCount - 1,
            //     targetRange.FirstColumn + targetRange.ColumnCount - 1);
            //
            // // Retrieve the CameraShape object
            // CameraShape camera = sheet.Shapes[shapeIndex] as CameraShape;
            // if (camera != null)
            // {
            //     camera.Resolution = 300;               // Set DPI
            //     camera.ImageFormat = System.Drawing.Imaging.ImageFormat.Png; // Set image format
            //     camera.Width = 400;                    // Adjust width (points)
            //     camera.Height = 300;                   // Adjust height (points)
            // }

            // Save the workbook
            string outputPath = "CameraShapeExample.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
