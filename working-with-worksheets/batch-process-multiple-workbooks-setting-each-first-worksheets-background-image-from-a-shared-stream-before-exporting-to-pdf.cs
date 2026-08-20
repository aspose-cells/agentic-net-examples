// Title: Batch convert Excel workbooks to PDF with a common worksheet background using Aspose.Cells for .NET
// Description: C# program that scans a folder for Excel files, loads a single image into memory, applies it as the background of the first worksheet in each workbook, and saves the result as PDF files. Includes error handling for missing files and image‑read failures, and uses Aspose.Cells PdfSaveOptions for the export.
// Keywords: Aspose.Cells background image | batch Excel to PDF conversion | apply worksheet background C# | shared image stream Aspose.Cells | PdfSaveOptions .NET | automate Excel PDF export
// Common Searches: how to add the same background image to multiple Excel sheets with Aspose.Cells | batch export Excel files to PDF with a common background in C# | set worksheet background from stream Aspose.Cells | convert a folder of workbooks to PDF using Aspose.Cells
// Developer Intent: Load one image once, set it as the background of the first worksheet in every workbook in a directory, and generate PDF versions of all workbooks.
// Use Cases: Nightly generation of client reports where the first page must display a corporate watermark. | Processing uploaded Excel templates, enforcing a compliance background on the first sheet, and archiving them as PDFs. | Creating a PDF catalog from a batch of product spreadsheets with a unified header graphic.
// AI Prompts: Show how to select a different background image for each workbook based on its filename using Aspose.Cells. | Give an example of configuring PdfSaveOptions to set page orientation, margins, and retain worksheet background images. | Explain how to parallelize the batch conversion while reusing the same background byte array.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering; // For PdfSaveOptions

// C# program that scans a folder for Excel files, loads a single image into memory, applies it as the background of the first worksheet in each workbook, and saves the result as PDF files. Includes error handling for missing files and image‑read failures, and uses Aspose.Cells PdfSaveOptions for the export.
class Program
{
    static void Main()
    {
        // Directory containing source Excel files
        string sourceDir = @"C:\InputWorkbooks";
        // Directory where PDF files will be saved
        string outputDir = @"C:\OutputPdfs";
        Directory.CreateDirectory(outputDir);

        // Path to the shared background image
        string backgroundImagePath = @"C:\SharedResources\background.png";

        // Load the background image once into a byte array (if the file exists)
        byte[] backgroundBytes = null;
        if (File.Exists(backgroundImagePath))
        {
            try
            {
                using (FileStream bgStream = new FileStream(backgroundImagePath, FileMode.Open, FileAccess.Read))
                {
                    backgroundBytes = new byte[bgStream.Length];
                    bgStream.Read(backgroundBytes, 0, backgroundBytes.Length);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Failed to read background image. {ex.Message}");
                backgroundBytes = null;
            }
        }
        else
        {
            Console.WriteLine("Warning: Background image not found. Workbooks will be saved without a background.");
        }

        // Get all Excel files in the source directory (supports .xlsx, .xls, .xlsm)
        string[] workbookFiles = Directory.GetFiles(sourceDir, "*.*", SearchOption.TopDirectoryOnly);
        foreach (string workbookPath in workbookFiles)
        {
            // Filter only Excel files
            string ext = Path.GetExtension(workbookPath).ToLowerInvariant();
            if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm")
                continue;

            // Ensure the workbook file still exists before loading
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Skipping missing file: {workbookPath}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Set background image on the first worksheet if available
                if (backgroundBytes != null && backgroundBytes.Length > 0)
                {
                    Worksheet firstSheet = workbook.Worksheets[0];
                    firstSheet.BackgroundImage = backgroundBytes;
                }

                // Prepare PDF save options (customizations can be added here)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Define output PDF path
                string pdfFileName = Path.GetFileNameWithoutExtension(workbookPath) + ".pdf";
                string pdfPath = Path.Combine(outputDir, pdfFileName);

                // Save the workbook as PDF
                workbook.Save(pdfPath, pdfOptions);

                Console.WriteLine($"Processed '{Path.GetFileName(workbookPath)}' -> '{pdfFileName}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{Path.GetFileName(workbookPath)}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch processing completed.");
    }
}
