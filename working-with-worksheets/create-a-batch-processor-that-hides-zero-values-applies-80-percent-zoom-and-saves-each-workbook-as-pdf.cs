// Title: Batch Convert Excel to PDF with Zero Hiding and 80% Zoom using Aspose.Cells for .NET
// Description: A C# console app that scans a folder for .xlsx files, loads each workbook with Aspose.Cells, disables zero display, sets the page‑setup zoom to 80 %, and saves every workbook as a PDF in a target directory. Ideal for automated, consistent PDF generation from multiple Excel files.
// Keywords: Aspose.Cells batch PDF conversion | hide zero values Excel | set worksheet zoom Aspose.Cells | C# convert multiple .xlsx to PDF | Aspose.Cells PDFSaveOptions | .NET Excel to PDF automation
// Common Searches: How to batch convert Excel files to PDF with Aspose.Cells | Hide zeros when exporting Excel to PDF in C# | Set 80% zoom for PDF export using Aspose.Cells | Process all .xlsx files in a folder and save as PDF | Aspose.Cells PDFSaveOptions for visible sheets only
// Developer Intent: Automatically process a directory of Excel workbooks, suppress zero values, apply an 80 % zoom setting, and generate a PDF for each file using Aspose.Cells for .NET.
// Use Cases: Create printable PDFs for financial statements where zero amounts should not appear. | Run a nightly job that archives client spreadsheets as PDFs with a uniform zoom level. | Prepare batch PDFs for regulatory filing, ensuring consistent page layout and hidden zero values.
// AI Prompts: Generate C# code that uses Aspose.Cells to batch convert all .xlsx files in a folder to PDF, hide zero values, and set page zoom to 80 %. | Explain how to extend the batch processor to include subfolders and customize PDF quality settings. | Show how to add robust error handling and logging to the Aspose.Cells PDF conversion loop.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace BatchPdfProcessor
{
    // A C# console app that scans a folder for .xlsx files, loads each workbook with Aspose.Cells, disables zero display, sets the page‑setup zoom to 80 %, and saves every workbook as a PDF in a target directory. Ideal for automated, consistent PDF generation from multiple Excel files.
    class Program
    {
        static void Main()
        {
            // Folder containing the source Excel workbooks
            string sourceFolder = @"C:\InputWorkbooks";

            // Folder where the resulting PDF files will be saved
            string outputFolder = @"C:\OutputPdfs";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all Excel files in the source folder (including subfolders if needed)
            string[] excelFiles = Directory.GetFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string excelPath in excelFiles)
            {
                // Load the workbook (uses the standard Workbook constructor)
                Workbook workbook = new Workbook(excelPath);

                // Process each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Hide zero values
                    sheet.DisplayZeros = false;

                    // Apply 80% zoom for printing/rendering
                    sheet.PageSetup.Zoom = 80;
                }

                // Prepare PDF save options (default options are sufficient for this task)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Ensure only visible sheets are exported (default behavior)
                    SheetSet = SheetSet.Visible
                };

                // Build the output PDF file name based on the source workbook name
                string pdfFileName = Path.GetFileNameWithoutExtension(excelPath) + ".pdf";
                string pdfPath = Path.Combine(outputFolder, pdfFileName);

                // Save the workbook as PDF using the provided save method
                workbook.Save(pdfPath, pdfOptions);

                Console.WriteLine($"Processed '{excelPath}' -> '{pdfPath}'");
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
