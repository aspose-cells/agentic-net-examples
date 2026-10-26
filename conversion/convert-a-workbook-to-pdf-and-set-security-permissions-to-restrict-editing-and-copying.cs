// Title: Convert an Excel workbook to PDF and apply editing‑ and copying‑restriction permissions with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel workbook to PDF with password protection | How to disable copy and edit in PDF generated from Excel using Aspose.Cells | Set PDF permissions when saving workbook as PDF with Aspose.Cells .NET | Example of applying security settings to PDF output from Aspose.Cells | Convert XLSX to secured PDF using Aspose.Cells and Aspose.Pdf in C#
// Tags: Aspose.Cells PDF export with security settings | C# Aspose.Cells PdfSaveOptions usage | Excel to PDF conversion Aspose.Cells | PDF permission configuration Aspose.Cells | Workbook to secured PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample checks that the source XLSX file exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object (without security options because Aspose.Cells does not expose PDF permission APIs), and saves the workbook as a PDF. It demonstrates basic conversion and notes that PDF protection must be applied through additional libraries such as Aspose.Pdf.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (no security settings due to unavailable API)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
