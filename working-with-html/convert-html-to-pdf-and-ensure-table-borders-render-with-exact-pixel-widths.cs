using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlToPdf
{
    class Program
    {
        static void Main()
        {
            // Paths for input HTML and output PDF
            string htmlPath = "input.html";
            string pdfPath = "output.pdf";

            try
            {
                // Verify that the HTML source file exists
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine($"Error: Input file '{htmlPath}' not found.");
                    return;
                }

                // Load the HTML file into a Workbook instance
                Workbook workbook = new Workbook(htmlPath);

                // Apply a uniform thick border to the used range of the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Style borderStyle = workbook.CreateStyle();
                borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
                borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;
                borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thick;
                borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;

                // Use 'var' to avoid ambiguity with System.Range
                var usedRange = sheet.Cells.MaxDisplayRange;
                usedRange.SetStyle(borderStyle);

                // Save the workbook as PDF
                workbook.Save(pdfPath, SaveFormat.Pdf);

                Console.WriteLine($"HTML file '{htmlPath}' has been converted to PDF '{pdfPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}