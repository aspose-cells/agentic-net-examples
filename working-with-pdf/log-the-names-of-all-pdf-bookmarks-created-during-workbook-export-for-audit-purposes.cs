// Title: Audit PDF Bookmarks When Exporting a Workbook to PDF with Aspose.Cells for .NET
// Description: This C# example creates a workbook, builds a hierarchical PdfBookmarkEntry structure, recursively logs each bookmark title, assigns the hierarchy to PdfSaveOptions.Bookmark, and saves the file as PDF, giving you a complete audit trail of every PDF bookmark generated during export.
// Keywords: Aspose.Cells | PDF bookmarks | C# | .NET | PdfBookmarkEntry | export workbook to PDF | bookmark audit | log bookmark names | hierarchical bookmarks | PDF conversion
// Common Searches: Aspose.Cells log PDF bookmark names C# | How to retrieve PDF bookmarks during workbook export | Audit PDF bookmark hierarchy Aspose.Cells .NET | Capture list of PDF bookmarks before saving | Get PdfBookmarkEntry titles in Aspose.Cells
// Developer Intent: Collect the titles of all PDF bookmarks created during a workbook‑to‑PDF conversion for auditing, validation, or compliance reporting.
// Use Cases: Print the full bookmark hierarchy to the console for quick verification. | Write bookmark titles to a log file or database to satisfy regulatory audit requirements. | Compare the generated bookmark list with an expected set to detect missing or extra entries before saving. | Integrate bookmark logging into a CI/CD pipeline to ensure consistent PDF navigation structures.
// AI Prompts: Generate C# code that traverses a PdfBookmarkEntry tree and writes each bookmark title to a text file for audit purposes. | Show how to capture PDF bookmark names after calling Workbook.Save with PdfSaveOptions in Aspose.Cells and store them in a JSON report. | Provide an example that uses StringBuilder to collect bookmark titles and includes timestamps for compliance logging.

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsBookmarkAudit
{
    // This C# example creates a workbook, builds a hierarchical PdfBookmarkEntry structure, recursively logs each bookmark title, assigns the hierarchy to PdfSaveOptions.Bookmark, and saves the file as PDF, giving you a complete audit trail of every PDF bookmark generated during export.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (contains a default sheet)
                Workbook workbook = new Workbook();

                // Remove the default worksheet to avoid duplicate name errors
                if (workbook.Worksheets.Count > 0)
                {
                    workbook.Worksheets.RemoveAt(0);
                }

                // Add three worksheets with unique names
                Worksheet sheet1 = workbook.Worksheets.Add("Sheet1");
                Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
                Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

                // Set values that will serve as bookmark destinations
                sheet1.Cells["A1"].PutValue("Content of Sheet1");
                sheet2.Cells["A1"].PutValue("Content of Sheet2");
                sheet3.Cells["A1"].PutValue("Content of Sheet3");

                // Build bookmark hierarchy
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

                // Attach sub‑bookmarks to the root
                rootBookmark.SubEntry.Add(subBookmark1);
                rootBookmark.SubEntry.Add(subBookmark2);

                // Log all bookmark names before exporting
                Console.WriteLine("PDF Bookmarks to be created:");
                LogBookmarkNames(rootBookmark, 0);

                // Configure PDF save options with the bookmark hierarchy
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    Bookmark = rootBookmark
                };

                // Save the workbook as PDF
                workbook.Save("output.pdf", pdfOptions);
                Console.WriteLine("PDF saved successfully as output.pdf");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Recursively traverses the PdfBookmarkEntry tree and writes each bookmark's Text
        static void LogBookmarkNames(PdfBookmarkEntry entry, int level)
        {
            if (entry == null) return;

            // Indent according to hierarchy level for readability
            string indent = new string(' ', level * 2);
            Console.WriteLine($"{indent}- {entry.Text}");

            // Process child entries if any
            if (entry.SubEntry != null)
            {
                foreach (PdfBookmarkEntry child in entry.SubEntry)
                {
                    LogBookmarkNames(child, level + 1);
                }
            }
        }
    }
}
