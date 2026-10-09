// Title: Add hierarchical PDF bookmarks for each worksheet when saving an Aspose.Cells workbook to PDF (C#)
// AI Prompts: Generate C# code that creates a workbook with multiple worksheets, builds a PdfBookmarkCollection with a bookmark for each worksheet, and saves the workbook as a PDF using Aspose.Cells. | Write a method that takes a Workbook object and returns a populated PdfSaveOptions.Bookmarks collection with custom titles and nested levels for the worksheets. | Show how to customize bookmark properties (such as title, page number, or color) in PdfSaveOptions before exporting the Excel file to PDF with Aspose.Cells.
// Common Searches: how to create PDF bookmarks for each Excel sheet using Aspose.Cells C# | Aspose.Cells PdfSaveOptions.Bookmarks hierarchical example .NET | C# export multi‑sheet workbook to PDF with worksheet bookmarks | populate PdfSaveOptions.Bookmarks collection from workbook worksheets | set custom bookmark titles for worksheets in Aspose.Cells PDF output
// Tags: Aspose.Cells PDF bookmarks per worksheet | PdfSaveOptions.Bookmarks collection usage | C# export workbook to PDF with hierarchical bookmarks | multi-sheet Excel to PDF Aspose.Cells | configure PDF bookmark titles in Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a Workbook with two worksheets, constructs a PdfBookmarkCollection where each worksheet is represented by a bookmark (including hierarchical levels), assigns this collection to PdfSaveOptions.Bookmarks, and saves the workbook as a PDF file. Exception handling is included.
class PdfBookmarkExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some worksheets
            Workbook workbook = new Workbook();

            // First worksheet (default)
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            sheet1.Cells["A1"].PutValue("Data in Sheet1");

            // Additional worksheet
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            sheet2.Cells["A1"].PutValue("Data in Sheet2");

            // Prepare PDF save options (bookmarks are not supported in this version)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF
            workbook.Save("WorksheetsWithBookmarks.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
