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
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape that will serve as the background
            // Parameters: upperLeftRow, upperLeftColumn, topOffset, leftOffset, width, height
            Shape background = sheet.Shapes.AddRectangle(0, 0, 0, 0, 500, 300);
            background.Name = "Background";

            // Send the background shape to the back (lowest Z‑order)
            // 0 = send to back, 1 = bring to front
            background.ToFrontOrBack(0);

            // Add another shape to verify that it appears above the background
            Shape foreground = sheet.Shapes.AddRectangle(2, 2, 20, 20, 200, 150);
            foreground.Name = "Foreground";

            // Ensure the output directory exists
            string outputPath = "ZOrderDemo.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}