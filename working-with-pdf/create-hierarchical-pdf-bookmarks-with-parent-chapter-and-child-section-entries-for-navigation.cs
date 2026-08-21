// Title: Create hierarchical PDF bookmarks (chapters & sections) with Aspose.Cells for .NET
// Description: Demonstrates how to build a workbook with chapter and section worksheets, construct a PdfBookmarkEntry hierarchy where each chapter contains child section bookmarks, attach the hierarchy to PdfSaveOptions, and export the workbook as a PDF with a multi‑level bookmark outline and document title.
// Keywords: Aspose.Cells PDF bookmarks | C# PdfBookmarkEntry hierarchy | nested PDF bookmarks Aspose.Cells | export Excel to PDF with bookmarks | chapter and section PDF outline | PdfSaveOptions Bookmark property | Aspose.Cells .NET example
// Common Searches: Aspose.Cells create nested PDF bookmarks C# | how to add chapter bookmarks to PDF using Aspose.Cells | PdfBookmarkEntry example for hierarchical bookmarks | export Excel workbook to PDF with outline in .NET | Aspose.Cells PDF bookmark hierarchy tutorial
// Developer Intent: Generate a PDF from an Excel workbook that includes parent chapter bookmarks and child section bookmarks for easy navigation.
// Use Cases: Convert a multi‑sheet workbook into an e‑book PDF where each sheet is a chapter and specific cells act as subsections. | Produce a structured PDF report with collapsible bookmark entries that mirror the workbook’s logical hierarchy. | Create a PDF document with a visible title and retained document structure for compliance or archival purposes.
// AI Prompts: Write C# code using Aspose.Cells to add parent chapter bookmarks and child section bookmarks to a PDF. | Explain the steps to build a PdfBookmarkEntry tree, assign it to PdfSaveOptions.Bookmark, and save a workbook as a PDF with nested bookmarks. | Show how to use a hidden root bookmark so top‑level chapters appear without an extra parent node in the PDF outline.

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to build a workbook with chapter and section worksheets, construct a PdfBookmarkEntry hierarchy where each chapter contains child section bookmarks, attach the hierarchy to PdfSaveOptions, and export the workbook as a PDF with a multi‑level bookmark outline and document title.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // -----------------------------------------------------------------
        // Prepare worksheets that will serve as chapters and sections
        // -----------------------------------------------------------------
        Worksheet chapter1 = workbook.Worksheets[0];
        chapter1.Name = "Chapter 1";
        chapter1.Cells["A1"].PutValue("Chapter 1 Title");
        chapter1.Cells["A2"].PutValue("Section 1.1");
        chapter1.Cells["A3"].PutValue("Section 1.2");

        Worksheet chapter2 = workbook.Worksheets.Add("Chapter 2");
        chapter2.Cells["A1"].PutValue("Chapter 2 Title");
        chapter2.Cells["A2"].PutValue("Section 2.1");
        chapter2.Cells["A3"].PutValue("Section 2.2");

        Worksheet chapter3 = workbook.Worksheets.Add("Chapter 3");
        chapter3.Cells["A1"].PutValue("Chapter 3 Title");
        chapter3.Cells["A2"].PutValue("Section 3.1");
        chapter3.Cells["A3"].PutValue("Section 3.2");

        // -----------------------------------------------------------------
        // Build hierarchical PDF bookmarks
        // -----------------------------------------------------------------
        // Root bookmark is hidden (Text = null) so its children appear at top level
        PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
        {
            Text = null,
            SubEntry = new ArrayList(),
            IsOpen = true
        };

        // Helper method to create a chapter bookmark with its sections
        PdfBookmarkEntry CreateChapter(string chapterTitle, Cell chapterCell,
                                      string[] sectionTitles, Cell[] sectionCells)
        {
            var chapterEntry = new PdfBookmarkEntry
            {
                Text = chapterTitle,
                Destination = chapterCell,
                SubEntry = new ArrayList(),
                IsOpen = true
            };

            for (int i = 0; i < sectionTitles.Length; i++)
            {
                var sectionEntry = new PdfBookmarkEntry
                {
                    Text = sectionTitles[i],
                    Destination = sectionCells[i]
                };
                chapterEntry.SubEntry.Add(sectionEntry);
            }

            return chapterEntry;
        }

        // Create chapter bookmarks
        var ch1Bookmark = CreateChapter(
            "Chapter 1",
            chapter1.Cells["A1"],
            new[] { "Section 1.1", "Section 1.2" },
            new[] { chapter1.Cells["A2"], chapter1.Cells["A3"] });

        var ch2Bookmark = CreateChapter(
            "Chapter 2",
            chapter2.Cells["A1"],
            new[] { "Section 2.1", "Section 2.2" },
            new[] { chapter2.Cells["A2"], chapter2.Cells["A3"] });

        var ch3Bookmark = CreateChapter(
            "Chapter 3",
            chapter3.Cells["A1"],
            new[] { "Section 3.1", "Section 3.2" },
            new[] { chapter3.Cells["A2"], chapter3.Cells["A3"] });

        // Attach chapter bookmarks to the root
        rootBookmark.SubEntry.Add(ch1Bookmark);
        rootBookmark.SubEntry.Add(ch2Bookmark);
        rootBookmark.SubEntry.Add(ch3Bookmark);

        // -----------------------------------------------------------------
        // Configure PDF save options with the bookmark hierarchy
        // -----------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Bookmark = rootBookmark,
            ExportDocumentStructure = true, // retain document structure
            DisplayDocTitle = true          // show document title in viewer
        };

        // Save the workbook as a PDF with hierarchical bookmarks
        workbook.Save("HierarchicalBookmarks.pdf", pdfOptions);
    }
}
