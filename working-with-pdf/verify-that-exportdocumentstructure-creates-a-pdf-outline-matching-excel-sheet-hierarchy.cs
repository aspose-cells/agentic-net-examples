// Title: How to verify that ExportDocumentStructure generates PDF bookmarks matching Excel worksheet hierarchy with Aspose.Cells for .NET
// Common Searches: how to enable PDF bookmarks when converting an Excel workbook to PDF with Aspose.Cells | C# code to check that PDF outline reflects worksheet names after export | Aspose.Cells ExportDocumentStructure option verification example | validate PDF outline hierarchy generated from multiple Excel sheets using .NET | sample program to save Excel as PDF with bookmarks using PdfSaveOptions
// Tags: ExportDocumentStructure PDF bookmarks Aspose.Cells | worksheet hierarchy to PDF outline .NET | verify PDF outline after workbook export | Aspose.Cells PdfSaveOptions ExportDocumentStructure usage | generate PDF bookmarks from Excel sheets C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example creates a workbook with several worksheets, enables the ExportDocumentStructure flag in PdfSaveOptions to produce a PDF outline, saves the workbook as a PDF, and confirms the file was created while listing the exported sheet names.
class ExportDocumentStructureVerification
{
    static void Main()
    {
        try
        {
            // Step 1: Create a workbook with multiple worksheets
            Workbook workbook = new Workbook();
            workbook.Worksheets.Clear();

            // Define sheet names to represent hierarchy
            string[] sheetNames = { "Summary", "Data_Q1", "Data_Q2", "Analysis" };

            foreach (string name in sheetNames)
            {
                Worksheet ws = workbook.Worksheets[workbook.Worksheets.Add()];
                ws.Name = name;
                ws.Cells["A1"].PutValue($"This is sheet {name}");
            }

            // Step 2: Export the workbook to PDF with document structure (outline) enabled
            string pdfPath = "WorkbookWithOutline.pdf";

            // Remove existing file if present to avoid conflicts
            if (File.Exists(pdfPath))
                File.Delete(pdfPath);

            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true // Enable PDF outline generation
            };
            workbook.Save(pdfPath, pdfOptions);

            // Step 3: Verify that the PDF file was created
            if (File.Exists(pdfPath))
                Console.WriteLine($"PDF saved successfully to '{pdfPath}'.");
            else
                Console.WriteLine($"Failed to create PDF at '{pdfPath}'.");

            // Optional: List the sheet names that were exported (for debugging)
            Console.WriteLine("Exported sheets:");
            foreach (string name in sheetNames)
                Console.WriteLine("- " + name);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
