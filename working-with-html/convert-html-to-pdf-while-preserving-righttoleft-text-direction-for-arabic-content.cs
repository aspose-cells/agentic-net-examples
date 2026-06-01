using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace HtmlToPdfRtlDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.html";
                const string outputPath = "output.pdf";

                // Verify that the input HTML file exists.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the HTML file into a workbook using LoadOptions for HTML format.
                var loadOptions = new LoadOptions(LoadFormat.Html);
                var workbook = new Workbook(inputPath, loadOptions);

                // Enable right‑to‑left display for every worksheet.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    sheet.DisplayRightToLeft = true;
                }

                // Set the default cell style to RightToLeft for the whole workbook.
                Style defaultStyle = workbook.DefaultStyle;
                defaultStyle.TextDirection = TextDirectionType.RightToLeft;
                workbook.DefaultStyle = defaultStyle;

                // Configure PDF save options with a font that supports Arabic glyphs.
                var pdfOptions = new PdfSaveOptions
                {
                    DefaultFont = "Arial"
                };

                // Save the workbook as a PDF file, preserving the RTL layout.
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF generated successfully: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}