using System;
using System.IO;
using Aspose.Cells;

class DisableBlankPagePdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a second (dummy) worksheet so the workbook has at least one visible sheet
            int dummyIndex = workbook.Worksheets.Add();

            // Hide the original first worksheet to simulate no printable content
            Worksheet firstSheet = workbook.Worksheets[0];
            firstSheet.IsVisible = false;

            // Configure PDF save options to suppress blank page generation
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OutputBlankPageWhenNothingToPrint = false
            };

            // Define the output file path (e.g., on the desktop)
            string outputPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "NoBlankPage.pdf");

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}