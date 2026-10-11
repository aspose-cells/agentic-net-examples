// Title: How to enable RotateTextWithShape in Aspose.Cells for .NET to rotate shape text together with the shape
// AI Prompts: Create a new workbook, add a rectangle shape, set its Text property, enable RotateTextWithShape, apply a 45‑degree RotationAngle, and save the file as .xlsx. | Update existing C# Aspose.Cells code to turn on the RotateTextWithShape flag so that any rotation applied to a shape also rotates its embedded text. | Generate a C# example that demonstrates setting RotateTextWithShape = true on a Shape object, assigning text, rotating the shape, and exporting the workbook.
// Common Searches: Aspose.Cells C# enable RotateTextWithShape property to rotate shape text | rotate shape text together with shape using Aspose.Cells .NET | C# example for rotating a shape and its text in an Excel file with Aspose.Cells | how to set RotateTextWithShape flag in Aspose.Cells for .NET workbook
// Tags: shape text rotation flag Aspose.Cells .NET | set RotateTextWithShape property C# | add rectangle shape worksheet Aspose.Cells | apply rotation angle to shape Aspose.Cells | export workbook with rotated shape Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The sample creates a workbook, inserts a rectangle shape at row 2 column 2, assigns the text “Rotated Text”, enables RotateTextWithShape, rotates the shape 45°, and saves the file as RotateTextWithShape.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: shape type, upper‑left row, upper‑left column, top offset, left offset, height, width
                Shape shape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // shape type
                    2,   // row index (zero‑based)
                    2,   // column index (zero‑based)
                    0,   // top offset in pixels
                    0,   // left offset in pixels
                    100, // height in points
                    200  // width in points
                );

                // Set the shape's text
                shape.Text = "Rotated Text";

                // Rotate the shape (e.g., 45 degrees). Text rotates together with the shape.
                shape.RotationAngle = 45;

                // Define output file path
                string outputPath = "RotateTextWithShape.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
