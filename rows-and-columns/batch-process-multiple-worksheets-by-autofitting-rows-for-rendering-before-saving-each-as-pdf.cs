// Title: Batch auto‑fit rows for each worksheet and save them as separate PDF files using Aspose.Cells for .NET
// Common Searches: C# auto‑fit rows in all worksheets before converting each sheet to PDF with Aspose.Cells | How to export every Excel worksheet to an individual PDF after adjusting row heights in .NET | Batch process an Excel workbook to generate per‑sheet PDFs with row auto‑fit using Aspose.Cells | Save each worksheet as a PDF file with proper row height using Aspose.Cells in C# | Create a temporary workbook for single‑sheet PDF export after AutoFitRows in Aspose.Cells
// Tags: auto‑fit rows Aspose.Cells PDF export | per‑sheet PDF generation Aspose.Cells | temporary workbook single sheet PDF Aspose.Cells | batch worksheet processing C# Aspose.Cells | row height adjustment before PDF conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, iterates through all worksheets, auto‑fits each sheet's rows, copies the sheet into a temporary workbook, and saves that workbook as an individual PDF named after the worksheet, handling missing files and per‑sheet errors.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file
            string sourceFile = @"C:\Input\Workbook.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourceFile))
            {
                Console.WriteLine($"Source file not found: {sourceFile}");
                return;
            }

            // Folder where individual PDF files will be saved
            string outputFolder = @"C:\Output\PdfSheets";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Load the workbook
            Workbook workbook = new Workbook(sourceFile);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Auto‑fit all rows in the current worksheet to ensure proper rendering
                    sheet.AutoFitRows();

                    // Create a temporary workbook containing only the current sheet
                    Workbook tempWb = new Workbook();
                    Worksheet tempSheet = tempWb.Worksheets[0];
                    tempSheet.Copy(sheet);               // Copy content from the source sheet
                    tempSheet.Name = sheet.Name;          // Preserve original sheet name

                    // Prepare PDF save options (customize if needed)
                    PdfSaveOptions pdfOptions = new PdfSaveOptions();

                    // Build the output PDF file name using the worksheet name
                    string pdfPath = Path.Combine(outputFolder, $"{sheet.Name}.pdf");

                    // Save the current worksheet as a separate PDF file
                    tempWb.Save(pdfPath, pdfOptions);
                }
                catch (Exception exSheet)
                {
                    Console.WriteLine($"Error processing sheet '{sheet.Name}': {exSheet.Message}");
                }
            }

            Console.WriteLine("All worksheets have been auto‑fitted and saved as individual PDFs.");
        }
        catch (Exception ex)
        {
            // Log the exception details for troubleshooting
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
