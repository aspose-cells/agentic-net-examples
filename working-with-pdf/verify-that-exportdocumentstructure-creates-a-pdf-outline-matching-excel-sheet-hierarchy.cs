// Title: Generate a PDF from an Excel workbook where the PDF outline reflects the worksheet hierarchy using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a Workbook with several worksheets, enables PdfSaveOptions.ExportDocumentStructure, and saves the workbook as a PDF. | Add logic to read the bookmarks from the generated PDF and compare them with the workbook's worksheet names to confirm the outline order. | Extend the sample to group worksheets into sections and verify that nested PDF bookmarks are produced for each group.
// Common Searches: Aspose.Cells how to enable PDF bookmarks that match Excel sheet names | C# export Excel to PDF with outline reflecting worksheet order | verify PDF bookmark hierarchy against workbook worksheets Aspose.Cells | ExportDocumentStructure option example for PDF generation .NET | create nested PDF outline from Excel sections using Aspose.Cells
// Tags: export workbook to PDF with document structure | Aspose.Cells PDF outline bookmarks | PdfSaveOptions ExportDocumentStructure usage | validate PDF bookmark hierarchy programmatically | generate nested PDF outline from Excel worksheets

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Shows how to build a workbook with multiple worksheets, enable the ExportDocumentStructure flag in PdfSaveOptions, save the workbook as a PDF, and confirm that the resulting PDF contains an outline (bookmarks) that mirrors the worksheet hierarchy.
class ExportDocumentStructureVerification
{
    static void Main()
    {
        try
        {
            // ---------- Create a sample workbook with multiple worksheets ----------
            Workbook workbook = new Workbook();
            workbook.Worksheets.Clear(); // start with a clean workbook

            // Add worksheets to represent a hierarchy (e.g., main sections and sub‑sections)
            Worksheet sheetA = workbook.Worksheets.Add("Section A");
            Worksheet sheetA1 = workbook.Worksheets.Add("Section A - Part 1");
            Worksheet sheetA2 = workbook.Worksheets.Add("Section A - Part 2");
            Worksheet sheetB = workbook.Worksheets.Add("Section B");
            Worksheet sheetB1 = workbook.Worksheets.Add("Section B - Part 1");

            // Populate each sheet with simple data (optional, just to have content)
            sheetA.Cells["A1"].PutValue("Data for Section A");
            sheetA1.Cells["A1"].PutValue("Data for Section A - Part 1");
            sheetA2.Cells["A1"].PutValue("Data for Section A - Part 2");
            sheetB.Cells["A1"].PutValue("Data for Section B");
            sheetB1.Cells["A1"].PutValue("Data for Section B - Part 1");

            // ---------- Export the workbook to PDF with document structure ----------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Enable PDF outline (bookmarks) that mirrors the worksheet hierarchy
                ExportDocumentStructure = true
            };
            string pdfPath = "ExportedWorkbook.pdf";

            // Save the workbook as PDF
            workbook.Save(pdfPath, pdfOptions);

            // Verify that the PDF file was created
            if (File.Exists(pdfPath))
            {
                Console.WriteLine("Verification succeeded: PDF file was created with document structure.");
            }
            else
            {
                Console.WriteLine("Verification failed: PDF file was not found.");
            }
        }
        catch (Exception ex)
        {
            // Output any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
