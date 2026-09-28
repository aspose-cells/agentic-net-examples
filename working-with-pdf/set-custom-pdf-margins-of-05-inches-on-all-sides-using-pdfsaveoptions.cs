// Title: How to set 0.5‑inch margins on all sides when converting an Excel workbook to PDF with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, applies 0.5‑inch margins to the worksheet via PageSetup, and saves it as a PDF using Aspose.Cells. | Show how to configure PdfSaveOptions in Aspose.Cells so that custom page margins are retained during Excel‑to‑PDF conversion. | Write a .NET snippet that validates the source Excel file, sets margin values in points, and exports the workbook to PDF while handling possible exceptions.
// Common Searches: Aspose.Cells .NET set half inch margins before PDF export | C# convert Excel to PDF with custom page margins using Aspose | How to change PDF page margins when saving workbook with Aspose.Cells | PdfSaveOptions margin configuration example in C# | Set page margins in points for Excel worksheet before PDF conversion
// Tags: Aspose.Cells PDF margin configuration C# | PageSetup custom margins Aspose.Cells | PdfSaveOptions preserve margins Aspose | Excel to PDF conversion half‑inch margins | C# workbook save as PDF custom margins

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example checks for the input Excel file, loads it into a Workbook, sets 0.5‑inch (36‑point) margins on the first worksheet via PageSetup, creates a PdfSaveOptions instance, and saves the workbook as a PDF while handling errors.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the existing Excel workbook
                Workbook workbook = new Workbook(inputPath);

                // Set 0.5‑inch margins on the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                PageSetup pageSetup = sheet.PageSetup;
                const double inchesToPoints = 72.0; // 1 inch = 72 points
                double marginPoints = 0.5 * inchesToPoints;
                pageSetup.LeftMargin = marginPoints;
                pageSetup.RightMargin = marginPoints;
                pageSetup.TopMargin = marginPoints;
                pageSetup.BottomMargin = marginPoints;

                // Create PDF save options (default options are sufficient)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as PDF with the specified margins
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
