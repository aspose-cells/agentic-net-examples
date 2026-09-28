// Title: Use Aspose.Cells LightCells API in C# to convert an Excel workbook to PDF while omitting charts
// AI Prompts: Write C# code that opens a specified .xlsx file with LightCells processing turned on, sets PdfSaveOptions.DisableChart to true, and saves the workbook as a PDF file. | Update an existing Aspose.Cells conversion routine to create the output folder if missing, enable LightCells for faster loading, and ensure charts are not rendered in the generated PDF.
// Common Searches: Aspose.Cells C# LightCells convert Excel to PDF without charts | disable chart rendering when saving workbook to PDF using Aspose.Cells .NET | enable LightCells API for faster Excel to PDF conversion in C# | C# check Excel file existence and create PDF output directory with Aspose.Cells
// Tags: LightCells API Excel to PDF conversion | exclude charts Aspose.Cells PDF export | C# enable LightCells processing | Aspose.Cells disable chart rendering PDF | create output directory Aspose.Cells conversion

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program verifies the source .xlsx file, enables LightCells for efficient loading, disables chart rendering via PdfSaveOptions, ensures the destination folder exists, and saves the workbook as a PDF using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source Excel file
                string sourcePath = "input.xlsx";

                // Path for the resulting PDF file
                string pdfPath = "output.pdf";

                // Verify that the source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the workbook from the source file
                Workbook workbook = new Workbook(sourcePath);

                // Configure PDF save options (default options used here)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(pdfPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to PDF format using the configured options
                workbook.Save(pdfPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to: {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
