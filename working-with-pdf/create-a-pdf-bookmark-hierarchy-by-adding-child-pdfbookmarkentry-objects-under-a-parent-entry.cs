using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class PdfBookmarkHierarchyDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and ensure it starts empty
            Workbook workbook = new Workbook();
            workbook.Worksheets.Clear(); // Remove the default sheet to avoid duplicate names

            // Add three worksheets with unique names
            workbook.Worksheets.Add("Sheet1");
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Define cells that will serve as bookmark destinations
            Cell cell1 = workbook.Worksheets[0].Cells["A1"];
            Cell cell2 = workbook.Worksheets[1].Cells["A1"];
            Cell cell3 = workbook.Worksheets[2].Cells["A1"];
            cell1.PutValue("Content of Sheet1");
            cell2.PutValue("Content of Sheet2");
            cell3.PutValue("Content of Sheet3");

            // Create the root PDF bookmark entry
            PdfBookmarkEntry root = new PdfBookmarkEntry
            {
                Text = "Workbook",
                Destination = cell1,
                IsOpen = true,
                SubEntry = new ArrayList()
            };

            // First-level child bookmark
            PdfBookmarkEntry child1 = new PdfBookmarkEntry
            {
                Text = "Section 1",
                Destination = cell2
            };

            // Second-level child (nested under child1)
            PdfBookmarkEntry subChild = new PdfBookmarkEntry
            {
                Text = "Subsection A",
                Destination = cell3
            };
            child1.SubEntry = new ArrayList { subChild };

            // Another first-level child bookmark
            PdfBookmarkEntry child2 = new PdfBookmarkEntry
            {
                Text = "Section 2",
                Destination = cell3,
                IsCollapse = true
            };

            // Assemble the hierarchy
            root.SubEntry.Add(child1);
            root.SubEntry.Add(child2);

            // Configure PDF save options with the bookmark hierarchy
            PdfSaveOptions options = new PdfSaveOptions
            {
                Bookmark = root,
                ExportDocumentStructure = true
            };

            // Save the workbook as a PDF file
            workbook.Save("WorkbookBookmarks.pdf", options);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}