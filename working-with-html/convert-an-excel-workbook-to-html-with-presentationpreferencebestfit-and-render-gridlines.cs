// Title: C# – Convert Excel to HTML with BestFit layout and gridlines using Aspose.Cells
// Description: Loads an .xlsx file (or creates a sample workbook), enables gridlines on the first worksheet, sets HtmlSaveOptions.PresentationPreference to BestFit and ExportGridLines to true, then saves the workbook as an HTML page that preserves column widths and shows gridlines.
// Keywords: Aspose.Cells | C# Excel to HTML | HtmlSaveOptions | PresentationPreference BestFit | ExportGridLines | gridlines visible | convert workbook to HTML | Aspose.Cells .NET | Excel HTML preview | HTML conversion Aspose
// Common Searches: Aspose.Cells export Excel to HTML with gridlines | HtmlSaveOptions PresentationPreference BestFit C# example | How to keep Excel column widths when converting to HTML | C# convert workbook to HTML showing gridlines | Aspose.Cells HTML conversion best fit layout
// Developer Intent: Generate an HTML file from an Excel workbook in C# with optimal column sizing and visible gridlines using Aspose.Cells.
// Use Cases: Render uploaded spreadsheets as web‑ready HTML previews that retain original column widths and gridlines. | Automate daily report publishing by converting Excel workbooks to HTML emails with a clean layout. | Provide an on‑demand Excel‑to‑HTML viewer in a web portal that shows exact worksheet structure.
// AI Prompts: Write C# code that opens an Excel file, sets HtmlSaveOptions.PresentationPreference = PresentationPreference.BestFit, enables ExportGridLines, and saves the result as HTML with Aspose.Cells. | Show how to ensure the first worksheet’s IsGridlinesVisible property is true before exporting to HTML. | Create a robust C# routine that converts multiple workbooks to HTML, creates missing output folders, and logs any conversion errors.

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelToHtmlConversion
{
    // Loads an .xlsx file (or creates a sample workbook), enables gridlines on the first worksheet, sets HtmlSaveOptions.PresentationPreference to BestFit and ExportGridLines to true, then saves the workbook as an HTML page that preserves column widths and shows gridlines.
    class Program
    {
        static void Main()
        {
            // Define input and output paths
            string inputPath = @"C:\Path\To\InputWorkbook.xlsx";
            string outputPath = @"C:\Path\To\OutputWorkbook.html";

            try
            {
                // Verify that the input file exists; if not, create a simple workbook as a placeholder
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file not found at '{inputPath}'. Creating a sample workbook.");
                    workbook = new Workbook();
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Cells["A1"].PutValue("Sample Data");
                }

                // Ensure gridlines are visible in the first worksheet
                Worksheet firstSheet = workbook.Worksheets[0];
                firstSheet.IsGridlinesVisible = true;

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    PresentationPreference = true, // Better layout
                    ExportGridLines = true          // Include gridlines in HTML
                };

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as HTML
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine("Workbook successfully converted to HTML with gridlines and presentation preference.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during conversion: {ex.Message}");
            }
        }
    }
}
