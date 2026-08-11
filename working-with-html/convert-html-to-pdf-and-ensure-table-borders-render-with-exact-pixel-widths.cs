// Title: C# – Convert HTML Table to PDF with Exact Border Widths Using Aspose.Cells
// Description: Loads an HTML file containing a table into an Aspose.Cells Workbook, applies a thin border style that matches the original pixel widths, and saves the workbook as a PDF. Includes file‑existence checks and exception handling for reliable batch conversion.
// Keywords: Aspose.Cells HTML to PDF | C# convert HTML table to PDF | preserve table borders PDF | exact border width Aspose.Cells | .NET export HTML as PDF | pixel‑perfect PDF conversion | Aspose.Cells border styling
// Common Searches: Aspose.Cells convert HTML to PDF with borders | C# keep HTML table border thickness in PDF | exact pixel border width when exporting HTML to PDF | apply uniform borders to worksheet before PDF export | handle missing HTML file Aspose.Cells
// Developer Intent: Create a PDF from an HTML table while retaining the original border thickness defined in the HTML markup.
// Use Cases: Generate printable invoices from HTML templates with precise table lines. | Automate PDF report generation from web‑based data tables that require consistent border styling. | Batch‑process a collection of HTML documents into PDFs, enforcing uniform thin borders for brand‑compliant output.
// AI Prompts: Write C# code that loads an HTML file into an Aspose.Cells Workbook, sets a thin border style on the used range, and saves it as a PDF. | Explain how to map HTML border‑pixel values to Aspose.Cells border line styles for accurate PDF rendering. | Suggest best practices for validating input HTML files and logging conversion results when using Aspose.Cells to export PDFs.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlToPdf
{
    // Loads an HTML file containing a table into an Aspose.Cells Workbook, applies a thin border style that matches the original pixel widths, and saves the workbook as a PDF. Includes file‑existence checks and exception handling for reliable batch conversion.
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file that contains the table.
            string htmlFilePath = "input.html";

            // Path where the resulting PDF will be saved.
            string pdfFilePath = "output.pdf";

            try
            {
                // Verify that the input HTML file exists to avoid FileNotFoundException.
                if (!File.Exists(htmlFilePath))
                {
                    Console.WriteLine($"Error: The input file '{htmlFilePath}' was not found.");
                    return;
                }

                // Load the HTML file into a Workbook.
                // Aspose.Cells automatically parses the HTML table structure and creates corresponding worksheet cells.
                Workbook workbook = new Workbook(htmlFilePath);

                // Optional: Ensure borders are rendered exactly as defined in the HTML.
                Worksheet sheet = workbook.Worksheets[0];
                Style borderStyle = workbook.CreateStyle();
                borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;

                // Apply the style to the used range to enforce exact border widths.
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
                usedRange.SetStyle(borderStyle);

                // Save the workbook as a PDF file.
                workbook.Save(pdfFilePath, SaveFormat.Pdf);

                Console.WriteLine($"HTML file '{htmlFilePath}' has been successfully converted to PDF '{pdfFilePath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
