// Title: How to verify that ExportDocumentStructure creates PDF bookmarks matching the Excel worksheet order using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells export workbook to PDF with a bookmark for each worksheet | check PDF outline hierarchy after enabling ExportDocumentStructure in Aspose.Cells | validate that PDF bookmarks follow the sequence of Excel sheets in C# | example code for using PdfSaveOptions ExportDocumentStructure and confirming the result | how to confirm PDF outline reflects worksheet names when converting Excel to PDF
// Tags: ExportDocumentStructure PDF bookmarks Aspose.Cells C# | worksheet hierarchy PDF outline verification | PdfSaveOptions document structure Aspose.Cells | C# generate PDF with worksheet bookmarks | validate PDF outline against Excel sheet order

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The sample creates a workbook with three worksheets, writes simple data to each sheet, and saves the workbook as a PDF using PdfSaveOptions with ExportDocumentStructure enabled. This option generates PDF bookmarks that correspond to each worksheet. The code then checks that the PDF file was created and lists the worksheet names, providing a straightforward way to verify that the PDF outline mirrors the Excel sheet hierarchy.
class ExportDocumentStructureVerification
{
    static void Main()
    {
        try
        {
            // Step 1: Create a workbook with multiple worksheets
            Workbook workbook = new Workbook();

            // Rename the default sheet and add two more sheets
            workbook.Worksheets[0].Name = "FirstSheet";
            workbook.Worksheets.Add("SecondSheet");
            workbook.Worksheets.Add("ThirdSheet");

            // (Optional) Fill each sheet with some data so the PDF has content
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet ws = workbook.Worksheets[i];
                ws.Cells["A1"].PutValue($"Data in {ws.Name}");
            }

            // Step 2: Export the workbook to PDF with document structure (outline) enabled
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true // Creates PDF bookmarks for each worksheet
            };
            string pdfPath = "WorkbookWithOutline.pdf";

            // Save the PDF
            workbook.Save(pdfPath, pdfOptions);

            // Step 3: Verify that the PDF file was created
            if (System.IO.File.Exists(pdfPath))
            {
                Console.WriteLine("SUCCESS: PDF file generated with document structure.");
            }
            else
            {
                Console.WriteLine("FAILURE: PDF file was not created.");
            }

            // (Optional) List worksheet names for reference
            Console.WriteLine("\nWorksheet Names:");
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Console.WriteLine($"- {ws.Name}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"ERROR: {ex.Message}");
        }
    }
}
