// Title: Convert Aspose.Cells Workbook to PDF with C# – Toolbar Hiding Not Supported
// Description: C# example that creates an Aspose.Cells workbook, adds sample data, configures PdfSaveOptions, ensures the output folder exists, and saves the workbook as a PDF. The code notes that current Aspose.Cells versions do not expose viewer‑preference settings such as hiding the PDF toolbar.
// Keywords: Aspose.Cells PDF conversion C# | save workbook as PDF | PdfSaveOptions viewer preferences | hide PDF toolbar Aspose.Cells | .NET Excel to PDF | Aspose.Cells export PDF
// Common Searches: Aspose.Cells hide toolbar when saving PDF | PdfSaveOptions hide PDF viewer toolbar C# | export Excel to PDF with custom viewer settings Aspose | C# convert workbook to PDF Aspose.Cells | viewer preferences not supported Aspose.Cells PDF
// Developer Intent: Generate a PDF from an Excel workbook using Aspose.Cells for .NET and understand that toolbar visibility cannot be set through PdfSaveOptions.
// Use Cases: Create a workbook programmatically, populate cells, and export it to PDF. | Automatically create the destination directory before writing the PDF file. | Identify the limitation that Aspose.Cells does not currently allow PDF viewer preferences such as toolbar visibility.
// AI Prompts: Write C# code with Aspose.Cells to convert a workbook to PDF and explain why toolbar hiding cannot be configured via PdfSaveOptions. | Suggest a post‑processing method or third‑party library to modify the generated PDF so the toolbar is hidden on open. | Describe how to monitor Aspose.Cells release notes for future support of PDF viewer preferences.

using System;
using System.IO;
using Aspose.Cells;

// C# example that creates an Aspose.Cells workbook, adds sample data, configures PdfSaveOptions, ensures the output folder exists, and saves the workbook as a PDF. The code notes that current Aspose.Cells versions do not expose viewer‑preference settings such as hiding the PDF toolbar.
class WorkbookToPdfWithHiddenToolbar
{
    static void Main()
    {
        try
        {
            // -----------------------------------------------------------------
            // 1. Create a new Excel workbook and add sample data
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data for PDF conversion");

            // -----------------------------------------------------------------
            // 2. Configure PDF save options
            // -----------------------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: In the current Aspose.Cells version, viewer preferences such as
            // hiding the toolbar are not exposed via PdfSaveOptions. The PDF will be
            // generated with default viewer settings.

            // -----------------------------------------------------------------
            // 3. Save the workbook as PDF
            // -----------------------------------------------------------------
            string pdfPath = "output.pdf";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(pdfPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine("Workbook has been saved to PDF.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
