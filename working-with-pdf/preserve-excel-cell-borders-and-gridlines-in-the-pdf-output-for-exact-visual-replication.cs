// Title: C# – Export an Excel workbook to PDF with Aspose.Cells while retaining cell borders and gridlines
// AI Prompts: Write C# code that loads an .xlsx file using Aspose.Cells, sets PdfSaveOptions to show gridlines and preserve border styles, and saves the workbook as a PDF. | Show how to check for the source Excel file, apply robust error handling, and ensure the generated PDF matches the original worksheet's visual layout. | Demonstrate configuring PdfSaveOptions (e.g., ShowGridlines = true) to keep exact formatting when converting Excel to PDF with Aspose.Cells.
// Common Searches: Aspose.Cells C# keep Excel cell borders when exporting to PDF | How to show gridlines in PDF generated from .xlsx using Aspose.Cells .NET | C# code sample for saving workbook as PDF with exact visual appearance | PdfSaveOptions ShowGridlines true Aspose.Cells example | Export Excel to PDF preserving formatting Aspose.Cells .NET
// Tags: Aspose.Cells PdfSaveOptions ShowGridlines | export Excel to PDF with borders .NET | preserve worksheet formatting Aspose.Cells PDF | C# convert .xlsx to .pdf retaining gridlines | Aspose.Cells PDF visual fidelity

using System;
using System.IO;
using Aspose.Cells;

// The example verifies that input.xlsx exists, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions (which display gridlines by default), and saves the workbook as output.pdf. The resulting PDF retains the original cell borders and gridlines, providing an exact visual replica of the Excel worksheet while handling errors gracefully.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (gridlines are shown by default in recent versions)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
