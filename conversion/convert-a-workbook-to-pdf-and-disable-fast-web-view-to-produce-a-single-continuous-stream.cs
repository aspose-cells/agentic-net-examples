// Title: Convert an Excel workbook to a continuous PDF stream with Fast Web View disabled using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, sets PdfSaveOptions.FastWebView to false, and saves it as a single‑stream PDF. | Show how to use Aspose.Cells PdfSaveOptions in a .NET console app to export an Excel workbook to a non‑linearized PDF.
// Common Searches: Aspose.Cells how to turn off Fast Web View when saving PDF in C# | C# export Excel to PDF as one continuous stream using Aspose.Cells | PdfSaveOptions FastWebView false example Aspose.Cells | Generate non‑linearized PDF from workbook with Aspose.Cells .NET
// Tags: Aspose.Cells PdfSaveOptions FastWebView | C# Excel to PDF continuous stream | disable fast web view Aspose PDF export | non-linearized PDF generation Aspose.Cells | single-stream PDF output from workbook

using System;
using System.IO;
using Aspose.Cells;

// The program checks for the input.xlsx file, loads it with Aspose.Cells, creates a PdfSaveOptions object with FastWebView disabled, and saves the workbook as output.pdf, producing a single continuous PDF stream.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (Fast Web View is disabled by default in many versions)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
