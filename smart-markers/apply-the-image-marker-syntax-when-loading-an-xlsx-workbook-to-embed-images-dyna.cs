using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsImageMarkerDemo
{
    class Program
    {
        static void Main()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;

            // Ensure template workbook exists
            string templatePath = Path.Combine(baseDir, "TemplateWithImageMarker.xlsx");
            if (!File.Exists(templatePath))
            {
                var wb = new Workbook();
                var ws = wb.Worksheets[0];
                ws.Cells["A1"].PutValue("&ImageData");
                wb.Save(templatePath);
            }

            // Ensure image file exists (1x1 transparent PNG)
            string imagePath = Path.Combine(baseDir, "DynamicImage.png");
            if (!File.Exists(imagePath))
            {
                byte[] pngBytes = Convert.FromBase64String(
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");
                File.WriteAllBytes(imagePath, pngBytes);
            }

            // Load the template workbook
            Workbook workbook = new Workbook(templatePath);

            // Read image bytes
            byte[] imageBytes = File.ReadAllBytes(imagePath);

            // Process smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("ImageData", imageBytes);
            designer.Process();

            // Save result
            string outputPath = Path.Combine(baseDir, "ResultWithEmbeddedImage.xlsx");
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }
}