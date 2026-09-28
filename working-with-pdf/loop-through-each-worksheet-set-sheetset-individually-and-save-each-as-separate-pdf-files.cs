// Title: Export each worksheet of an Excel workbook to a separate PDF file with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that iterates through all worksheets in a Workbook and saves each one as an individual PDF using PdfSaveOptions.SheetSet. | Show how to create an output folder beside the source Excel file and name each PDF after its worksheet name. | Demonstrate configuring PdfSaveOptions.SheetSet to a single sheet index so only that sheet is exported to PDF. | Provide robust error handling for a missing source file and ensure PDFs are written to a subdirectory.
// Common Searches: Aspose.Cells C# export each Excel sheet to its own PDF file | How to use PdfSaveOptions SheetSet to save a single worksheet as PDF in .NET | Create a folder and save worksheet PDFs with names matching sheet titles using Aspose.Cells | Loop through workbook worksheets and generate separate PDF documents in C# | Save Excel worksheets as individual PDFs with Aspose.Cells and handle missing file errors
// Tags: Aspose.Cells PdfSaveOptions SheetSet per worksheet | C# export Excel worksheets to separate PDF files | create output directory for PDF exports Aspose.Cells | save individual worksheet as PDF Aspose.Cells .NET | loop through workbook worksheets PDF generation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads a workbook, creates a "PdfExports" subfolder, iterates each worksheet, sets PdfSaveOptions.SheetSet to the current sheet index, and saves each sheet as a PDF named after the worksheet.
class ExportWorksheetsToPdf
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file
            string excelPath = @"C:\Data\Workbook.xlsx";

            // Verify that the source file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: The file \"{excelPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Determine output directory (same folder as source file)
            string sourceDir = Path.GetDirectoryName(excelPath) ?? ".";
            string outputDir = Path.Combine(sourceDir, "PdfExports");
            Directory.CreateDirectory(outputDir);

            // Export each worksheet to a separate PDF file
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Configure PDF save options for the current sheet only
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Use SheetSet to specify a single sheet by its index
                    SheetSet = new SheetSet(sheet.Index, sheet.Index)
                };

                // Build the output PDF file name (e.g., Sheet1.pdf)
                string pdfFileName = Path.Combine(outputDir, $"{sheet.Name}.pdf");

                // Save the workbook (only the current sheet) as a PDF file
                workbook.Save(pdfFileName, pdfOptions);
            }

            Console.WriteLine("All worksheets have been exported to separate PDF files.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
