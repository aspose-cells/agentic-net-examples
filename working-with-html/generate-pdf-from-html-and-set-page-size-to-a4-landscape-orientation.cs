using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlToPdfDemo
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                const string inputHtmlPath = "input.html";
                const string outputPdfPath = "output.pdf";

                // Verify that the HTML input file exists to avoid FileNotFoundException
                if (!File.Exists(inputHtmlPath))
                {
                    Console.WriteLine($"Error: The file \"{inputHtmlPath}\" was not found.");
                    return;
                }

                // Load the HTML file into a workbook using LoadOptions for HTML format
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
                Workbook workbook = new Workbook(inputHtmlPath, loadOptions);

                // Configure page settings (A4 size, landscape orientation)
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
                worksheet.PageSetup.Orientation = PageOrientationType.Landscape;

                // Save the workbook as a PDF file
                workbook.Save(outputPdfPath, SaveFormat.Pdf);

                Console.WriteLine("HTML has been successfully converted to PDF with A4 landscape orientation.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}