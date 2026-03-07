using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWordArtGradientToHtml
{
    class Program
    {
        static void Main()
        {
            // Determine the full path of the source Excel file relative to the executable directory
            string sourceFileName = "WordArtGradient.xlsx";
            string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), sourceFileName);
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "WordArtGradient.html");

            // If the source file does not exist, create a simple workbook as a placeholder
            if (!File.Exists(sourcePath))
            {
                var wb = new Workbook();
                var ws = wb.Worksheets[0];
                ws.Name = "Sample";
                ws.Cells["A1"].PutValue("Placeholder workbook created because the source file was missing.");
                wb.Save(sourcePath);
            }

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Set HTML save options (optional customizations can be added here)
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Save as HTML
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Conversion completed. HTML saved to: {outputPath}");
        }
    }
}