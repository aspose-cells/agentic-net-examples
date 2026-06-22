using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace InsertSvgExampleApp
{
    class InsertSvgExample
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Load SVG file if it exists
                string svgPath = "logo.svg";
                byte[] svgData = null;
                if (File.Exists(svgPath))
                {
                    svgData = File.ReadAllBytes(svgPath);
                }
                else
                {
                    Console.WriteLine($"SVG file not found: {svgPath}");
                }

                if (svgData != null)
                {
                    // Insert SVG into range E5:F6 (rows 4‑5, columns 4‑5)
                    // Width and height are required parameters; imageData can be null
                    int width = 200;   // desired width in pixels
                    int height = 200;  // desired height in pixels

                    ShapeCollection shapes = worksheet.Shapes;
                    // AddSvg(upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn, width, height, svgData, imageData)
                    shapes.AddSvg(4, 4, 5, 5, width, height, svgData, null);
                }

                // Save the workbook
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}