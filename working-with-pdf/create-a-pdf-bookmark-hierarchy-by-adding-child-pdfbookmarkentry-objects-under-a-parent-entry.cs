// Title: Create hierarchical PDF bookmarks with child and nested PdfBookmarkEntry objects using Aspose.Cells for .NET
// Description: This example builds a multi‑level PDF outline by defining a root PdfBookmarkEntry, adding chapter bookmarks as children, and nesting a sub‑section under one chapter via the SubEntry collections. The hierarchy is assigned to PdfSaveOptions.Bookmark, ExportDocumentStructure is enabled for accessibility, and the workbook is saved as a PDF with clickable bookmarks.
// Keywords: Aspose.Cells PDF bookmarks | PdfBookmarkEntry C# | nested PDF bookmarks .NET | PdfSaveOptions bookmark hierarchy | export Excel to PDF with outline | C# PDF bookmark example | Aspose.Cells PDF outline
// Common Searches: Aspose.Cells add child PDF bookmark .NET | Create nested PDF bookmarks from Excel worksheets | PdfSaveOptions Bookmark hierarchy example | Export workbook to PDF with multi‑level bookmarks | C# Aspose.Cells PDF outline accessibility
// Developer Intent: Generate a PDF from an Excel workbook that contains a structured, multi‑level bookmark outline.
// Use Cases: Produce a PDF report with a clickable table of contents linking to cover and chapter sheets. | Create accessible PDFs where each worksheet appears as a top‑level bookmark and sections within a sheet are nested bookmarks. | Automate e‑book generation with hierarchical bookmarks for chapters and sub‑sections.
// AI Prompts: Show how to build a PDF bookmark tree with child and nested PdfBookmarkEntry objects in Aspose.Cells for .NET. | Provide C# code that dynamically creates PDF bookmarks from a list of worksheet names and cell references. | Explain the effect of ExportDocumentStructure on PDF accessibility when bookmarks are attached.

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfBookmarkDemo
{
    // This example builds a multi‑level PDF outline by defining a root PdfBookmarkEntry, adding chapter bookmarks as children, and nesting a sub‑section under one chapter via the SubEntry collections. The hierarchy is assigned to PdfSaveOptions.Bookmark, ExportDocumentStructure is enabled for accessibility, and the workbook is saved as a PDF with clickable bookmarks.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add three worksheets
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Cover";
            Worksheet sheet2 = workbook.Worksheets.Add("Chapter1");
            Worksheet sheet3 = workbook.Worksheets.Add("Chapter2");

            // Put some sample data that will serve as bookmark destinations
            sheet1.Cells["A1"].PutValue("Cover Page");
            sheet2.Cells["A1"].PutValue("Chapter 1 Content");
            sheet3.Cells["A1"].PutValue("Chapter 2 Content");

            // ---------- Create PDF bookmark hierarchy ----------
            // Root bookmark (will appear as the top level entry in the PDF)
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = "Document Outline",
                Destination = sheet1.Cells["A1"],
                IsOpen = true,
                SubEntry = new ArrayList()
            };

            // First child bookmark
            PdfBookmarkEntry chapter1Bookmark = new PdfBookmarkEntry
            {
                Text = "Chapter 1",
                Destination = sheet2.Cells["A1"]
            };

            // Second child bookmark
            PdfBookmarkEntry chapter2Bookmark = new PdfBookmarkEntry
            {
                Text = "Chapter 2",
                Destination = sheet3.Cells["A1"],
                SubEntry = new ArrayList()
            };

            // Add a nested sub‑bookmark under Chapter 2
            PdfBookmarkEntry subSectionBookmark = new PdfBookmarkEntry
            {
                Text = "Section 2.1",
                Destination = sheet3.Cells["A1"] // using same cell for demo; replace with actual cell as needed
            };
            chapter2Bookmark.SubEntry.Add(subSectionBookmark);

            // Assemble the hierarchy
            rootBookmark.SubEntry.Add(chapter1Bookmark);
            rootBookmark.SubEntry.Add(chapter2Bookmark);

            // ---------- Configure PDF save options ----------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Attach the bookmark hierarchy to the PDF
                Bookmark = rootBookmark,
                // Optional: keep the document structure for better accessibility
                ExportDocumentStructure = true
            };

            // ---------- Save the workbook as PDF ----------
            workbook.Save("DocumentWithBookmarks.pdf", pdfOptions);

            Console.WriteLine("PDF with bookmark hierarchy created successfully.");
        }
    }
}
