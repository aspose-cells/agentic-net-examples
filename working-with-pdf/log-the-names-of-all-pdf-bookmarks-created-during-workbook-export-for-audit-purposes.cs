using System;
using System.Collections;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (remove the default sheet to avoid name conflicts)
            Workbook workbook = new Workbook();
            workbook.Worksheets.Clear();

            // Add worksheets and set sample data
            Worksheet sheet1 = workbook.Worksheets.Add("Sheet1");
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

            sheet1.Cells["A1"].PutValue("Content of Sheet1");
            sheet2.Cells["A1"].PutValue("Content of Sheet2");
            sheet3.Cells["A1"].PutValue("Content of Sheet3");

            // Build PDF bookmark hierarchy
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = "Root",
                Destination = sheet1.Cells["A1"],
                IsOpen = true,
                SubEntry = new ArrayList()
            };

            PdfBookmarkEntry subBookmark1 = new PdfBookmarkEntry
            {
                Text = "Sheet2",
                Destination = sheet2.Cells["A1"]
            };

            PdfBookmarkEntry subBookmark2 = new PdfBookmarkEntry
            {
                Text = "Sheet3",
                Destination = sheet3.Cells["A1"]
            };

            rootBookmark.SubEntry.Add(subBookmark1);
            rootBookmark.SubEntry.Add(subBookmark2);

            // Log all bookmark names for audit
            LogBookmarks(rootBookmark);

            // Save workbook to PDF with the defined bookmarks
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = rootBookmark
            };

            string outputPath = "output.pdf";
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Recursive method to traverse and log bookmark names
    static void LogBookmarks(PdfBookmarkEntry entry, string prefix = "")
    {
        if (!string.IsNullOrEmpty(entry.Text))
        {
            Console.WriteLine($"{prefix}Bookmark: {entry.Text}");
        }

        if (entry.SubEntry != null)
        {
            foreach (PdfBookmarkEntry child in entry.SubEntry)
            {
                LogBookmarks(child, prefix + "  ");
            }
        }
    }
}