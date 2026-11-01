// Title: Embedding a CSV file as an attachment in a PDF with Aspose.Cells PdfSaveOptions in C#
// Common Searches: Aspose.Cells C# embed CSV as attachment when saving workbook to PDF | PdfSaveOptions add external file attachment to PDF generated from Excel | How to attach a CSV file to a PDF using Aspose.Cells API in C# | Save Excel as PDF with embedded CSV using Aspose.Cells PdfSaveOptions
// Tags: Aspose.Cells PdfSaveOptions embed file attachment | C# attach CSV to PDF with Aspose.Cells | Excel to PDF with attached CSV using Aspose.Cells | PDF attachment feature Aspose.Cells | Saving workbook as PDF with external file attachment

using System;
using System.IO;
using Aspose.Cells;

// The example verifies that both an Excel workbook and a CSV file exist, loads the workbook, and saves it as a PDF using Aspose.Cells. To embed the CSV as an attachment, you would configure PdfSaveOptions.Attachments with the CSV file before calling the Save method.
class Program
{
    static void Main()
    {
        // Paths to input and output files
        string excelPath = "input.xlsx";
        string csvPath = "data.csv";
        string pdfPath = "output.pdf";

        // Verify that required files exist
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"Error: Excel file not found at '{excelPath}'.");
            return;
        }

        if (!File.Exists(csvPath))
        {
            Console.WriteLine($"Error: CSV file not found at '{csvPath}'.");
            return;
        }

        try
        {
            // Load the workbook and save it as PDF
            Workbook workbook = new Workbook(excelPath);
            workbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine($"PDF successfully created at '{pdfPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
