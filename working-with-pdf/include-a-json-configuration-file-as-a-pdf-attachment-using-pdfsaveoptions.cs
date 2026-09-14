// Title: Save an Excel workbook as PDF while embedding a JSON configuration in a hidden worksheet using Aspose.Cells PdfSaveOptions (C#)
// AI Prompts: Write C# code that reads a JSON file, stores its content in a hidden worksheet, and saves the workbook as a PDF with Aspose.Cells PdfSaveOptions. | Show how to conditionally add auxiliary JSON data to an Excel file before converting it to PDF using Aspose.Cells in .NET.
// Common Searches: how to add a hidden worksheet with JSON data before exporting to PDF using Aspose.Cells C# | Aspose.Cells PdfSaveOptions embed configuration file in Excel then convert to PDF | C# save workbook as PDF while keeping JSON config hidden in the file | Aspose.Cells include auxiliary data in PDF output without displaying it in Excel | read json file and hide it in Excel sheet then generate PDF with Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions hidden worksheet | C# embed json in Excel hidden sheet | Aspose.Cells export workbook to PDF with auxiliary data | conditional json embedding Aspose.Cells | Aspose.Cells PDF generation with hidden configuration

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfExample
{
    // The example checks for a config.json file, reads its bytes, and if present creates a hidden worksheet named 'JsonData' containing the JSON text. Sample content is added to the first sheet, PdfSaveOptions are configured, and the workbook is saved as output.pdf.
    class Program
    {
        static void Main()
        {
            try
            {
                const string jsonFile = "config.json";
                byte[] jsonData = null;

                // Load JSON file if it exists
                if (File.Exists(jsonFile))
                {
                    jsonData = File.ReadAllBytes(jsonFile);
                }
                else
                {
                    Console.WriteLine($"Warning: '{jsonFile}' not found. Continuing without attachment.");
                }

                // Create a new workbook and add sample content
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("PDF generated with a JSON attachment.");

                // If JSON data is available, store it in a hidden worksheet (alternative to attachment)
                if (jsonData != null)
                {
                    Worksheet hiddenSheet = workbook.Worksheets.Add("JsonData");
                    hiddenSheet.IsVisible = false;
                    string jsonText = System.Text.Encoding.UTF8.GetString(jsonData);
                    hiddenSheet.Cells["A1"].PutValue(jsonText);
                }

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as a PDF
                const string outputPdf = "output.pdf";
                workbook.Save(outputPdf, pdfOptions);
                Console.WriteLine($"PDF saved successfully to '{outputPdf}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
