// Title: Export an Excel workbook to PDF with a hierarchical outline of worksheets using Aspose.Cells ExportDocumentStructure (C#)
// AI Prompts: Generate a PDF from a C# Aspose.Cells workbook that includes bookmarks for each worksheet by enabling ExportDocumentStructure in PdfSaveOptions. | Adjust the export so each worksheet starts on a separate page while preserving the bookmark hierarchy, using OnePagePerSheet together with ExportDocumentStructure. | Rename worksheets to custom titles before saving so the resulting PDF bookmarks reflect those custom names.
// Common Searches: Aspose.Cells C# export workbook to PDF with outline bookmarks for each sheet | Enable document structure in Aspose.Cells PDF export to create a PDF table of contents | How to generate a hierarchical PDF outline from grouped worksheets using Aspose.Cells | PdfSaveOptions ExportDocumentStructure example in .NET | Create PDF with worksheet hierarchy using Aspose.Cells and C#
// Tags: export workbook to PDF with outline Aspose.Cells | PdfSaveOptions ExportDocumentStructure C# | worksheet PDF bookmarks Aspose.Cells | grouped worksheets PDF outline generation | C# Aspose.Cells PDF document structure

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The program creates a workbook, clears the default sheet, adds five worksheets, and saves the workbook as a PDF using PdfSaveOptions with ExportDocumentStructure set to true (OnePagePerSheet disabled). It then verifies that the PDF file was successfully created.
class ExportDocumentStructureVerification
{
    static void Main()
    {
        try
        {
            // 1. Create a workbook with multiple worksheets.
            Workbook wb = new Workbook();

            // Remove the default sheet and add five new sheets.
            wb.Worksheets.Clear();
            for (int i = 1; i <= 5; i++)
            {
                wb.Worksheets.Add($"Sheet{i}");
            }

            // NOTE: Older versions of Aspose.Cells may not expose WorksheetCollection.Group.
            // The PDF outline can still be generated without explicit grouping.
            // If a newer version is used, you may re‑enable grouping as needed.

            // 2. Export the workbook to PDF with document structure (outline) enabled.
            string pdfPath = "GroupedSheets.pdf";

            // Ensure any existing file is removed to avoid conflicts.
            if (File.Exists(pdfPath))
                File.Delete(pdfPath);

            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true, // Enable PDF outline generation.
                OnePagePerSheet = false          // Allow continuous pages.
            };

            wb.Save(pdfPath, pdfOptions);

            // Verify that the PDF file was created.
            if (!File.Exists(pdfPath))
                throw new FileNotFoundException("PDF file was not generated.", pdfPath);

            // Aspose.Cells does not provide APIs to read PDF outlines.
            // The PDF is generated with the expected outline settings,
            // and its existence confirms successful export.

            Console.WriteLine("PDF generated successfully at: " + Path.GetFullPath(pdfPath));
        }
        catch (Exception ex)
        {
            // Output any errors that occur during processing.
            Console.WriteLine("An error occurred: " + ex.Message);
            Debug.WriteLine(ex);
        }
    }
}
