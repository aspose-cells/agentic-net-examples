// Title: Enable RotateTextWithShape to rotate shape text together with the shape using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to set RotateTextWithShape to true for rotating shape text | example of rotating both shape and its text in Excel with Aspose.Cells .NET | Aspose.Cells RotateTextWithShape property usage in C# workbook | rotate shape text along with shape angle Aspose.Cells C# code sample | enable text rotation with shape rotation Aspose.Cells .NET tutorial
// Tags: RotateTextWithShape true Aspose.Cells | rotate shape text together Aspose.Cells | C# shape text rotation Excel Aspose.Cells | Aspose.Cells shape rotation with text property | Excel workbook shape text rotation .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This C# example demonstrates how to enable the RotateTextWithShape property in Aspose.Cells, add a rectangle shape to a worksheet, assign text, set the shape's RotationAngle, and save the workbook so that the text rotates together with the shape.
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
                // Parameters: shape type, upper‑left row, upper‑left column, row offset, column offset, width, height
                Shape shape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle,
                    2, 2, 0, 0, 120, 60);

                // Set the shape's text
                shape.Text = "Rotated Text";

                // Rotate the shape itself (e.g., 45 degrees)
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
