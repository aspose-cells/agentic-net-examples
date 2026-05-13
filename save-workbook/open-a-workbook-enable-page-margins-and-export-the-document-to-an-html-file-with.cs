using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook from disk
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Enable page margins by setting them on the first worksheet's PageSetup
            Worksheet sheet = workbook.Worksheets[0];
            // Margins are specified in points (1 point = 1/72 inch)
            sheet.PageSetup.TopMargin = 72;      // 1 inch
            sheet.PageSetup.BottomMargin = 72;   // 1 inch
            sheet.PageSetup.LeftMargin = 72;     // 1 inch
            sheet.PageSetup.RightMargin = 72;    // 1 inch

            // Configure HTML save options to embed images as Base64 strings
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = true
            };

            // Save the workbook as an HTML file with the specified options
            string outputPath = "output.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook exported to HTML with embedded images at: {outputPath}");
        }
    }
}