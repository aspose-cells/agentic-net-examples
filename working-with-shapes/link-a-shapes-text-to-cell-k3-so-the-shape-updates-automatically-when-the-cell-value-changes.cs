using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape.
            // Parameters: upper‑left row, upper‑left column,
            // upper‑left row offset (pixels), upper‑left column offset (pixels),
            // width (pixels), height (pixels)
            Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 50);

            // Link the shape's value to cell K3.
            shape.LinkedCell = "$K$3";

            // Define output file path
            string outputPath = "LinkedShape.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}