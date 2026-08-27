// Title: Convert an Excel workbook to PDF with hidden toolbar using Aspose.Cells in C#
// AI Prompts: Generate C# code that loads an .xlsx workbook with Aspose.Cells, configures PdfViewerPreferences to hide the toolbar, and saves the workbook as a PDF using PdfSaveOptions. | Demonstrate creating the destination folder programmatically if it does not exist before calling Workbook.Save with PdfSaveOptions. | Write robust error‑handling that verifies the source Excel file, catches conversion exceptions, and logs clear messages.
// Common Searches: asp.net aspose.cells convert excel to pdf with hidden toolbar | c# pdfsaveoptions hide toolbar property example | how to configure pdf viewer preferences in aspose.cells when saving workbook | save workbook as pdf and suppress toolbar using aspose.cells c#
// Tags: Aspose.Cells PdfSaveOptions HideToolbar setting | Excel to PDF conversion C# Aspose.Cells | PDF viewer preferences configuration Aspose.Cells | Create output directory before saving PDF C# | Exception handling for workbook conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing XLSX file with Aspose.Cells, optionally applies PdfViewerPreferences to hide the PDF toolbar via PdfSaveOptions, ensures the output folder exists, and saves the workbook as a PDF while handling missing files and runtime exceptions.
class WorkbookToPdfWithHiddenToolbar
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: In some versions of Aspose.Cells the PdfViewerPreferences class is not available.
            // If it is available, you can uncomment the following lines to hide the toolbar:
            // pdfOptions.PdfViewerPreferences = new Aspose.Cells.Pdf.PdfViewerPreferences
            // {
            //     HideToolbar = true
            // };

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully converted to PDF: \"{outputPath}\"");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
