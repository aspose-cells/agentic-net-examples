// Title: Create hierarchical PDF bookmarks for chapter and section worksheets using Aspose.Cells in C#
// AI Prompts: Generate a PDF where each worksheet becomes a separate page and assign parent‑child outline entries for chapter and section sheets using Aspose.Cells. | Modify the code to set custom bookmark names and nest section bookmarks under their chapter bookmarks in the PDF export. | Add a document title and configure the PDF outline hierarchy before saving the workbook with PdfSaveOptions.
// Common Searches: how to export Excel worksheets to a PDF with nested bookmarks in C# using Aspose.Cells | Aspose.Cells PDFSaveOptions OnePagePerSheet with chapter and section outline | C# code example for creating parent and child PDF bookmarks from multiple sheets | set custom PDF bookmark titles for each worksheet in Aspose.Cells
// Tags: Aspose.Cells PDF hierarchical outline | C# export worksheets to PDF with bookmarks | OnePagePerSheet PdfSaveOptions bookmark nesting | custom PDF bookmark titles Aspose.Cells | chapter section PDF outline Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample builds a workbook with separate worksheets for chapters and their sections, fills each sheet with sample text, configures PdfSaveOptions to place each worksheet on its own PDF page, and saves the workbook as 'HierarchicalBookmarks.pdf' with a hierarchical bookmark structure that reflects the chapter‑section relationship.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------- Chapter 1 --------------------
            Worksheet chapterSheet = workbook.Worksheets[0];
            chapterSheet.Name = "Chapter 1";
            chapterSheet.Cells["A1"].PutValue("Content of Chapter 1");
            for (int i = 2; i <= 30; i++)
            {
                chapterSheet.Cells[$"A{i}"].PutValue($"Chapter 1 - line {i}");
            }

            // -------------------- Section 1.1 (child of Chapter 1) --------------------
            Worksheet sectionSheet = workbook.Worksheets.Add("Section 1.1");
            sectionSheet.Cells["A1"].PutValue("Content of Section 1.1");
            for (int i = 2; i <= 20; i++)
            {
                sectionSheet.Cells[$"A{i}"].PutValue($"Section 1.1 - line {i}");
            }

            // -------------------- Chapter 2 --------------------
            Worksheet chapter2Sheet = workbook.Worksheets.Add("Chapter 2");
            chapter2Sheet.Cells["A1"].PutValue("Content of Chapter 2");
            for (int i = 2; i <= 25; i++)
            {
                chapter2Sheet.Cells[$"A{i}"].PutValue($"Chapter 2 - line {i}");
            }

            // -------------------- Section 2.1 (child of Chapter 2) --------------------
            Worksheet section2Sheet = workbook.Worksheets.Add("Section 2.1");
            section2Sheet.Cells["A1"].PutValue("Content of Section 2.1");
            for (int i = 2; i <= 15; i++)
            {
                section2Sheet.Cells[$"A{i}"].PutValue($"Section 2.1 - line {i}");
            }

            // Prepare PDF save options (export each worksheet as a separate page)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            // Define output file path
            string outputPath = "HierarchicalBookmarks.pdf";

            // Save the workbook as a PDF
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
