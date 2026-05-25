using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Alias to avoid ambiguity with System.Range
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // ------------------------------------------------------------
            // Sample data and named ranges – in real usage the workbook
            // would already contain the data and named ranges.
            // ------------------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Populate some cells
            sheet.Cells["A1"].PutValue("First range content");
            sheet.Cells["B2"].PutValue("Second range content");

            // Define named ranges
            sheet.Cells.CreateRange("A1:A1").Name = "FirstRange";
            sheet.Cells.CreateRange("B2:B2").Name = "SecondRange";

            // ------------------------------------------------------------
            // Build PDF bookmark hierarchy based on named ranges
            // ------------------------------------------------------------
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = "Named Ranges",   // Root title
                IsOpen = true,           // Expand root by default
                SubEntry = new ArrayList()
            };

            // Iterate through all defined names in the workbook
            foreach (Name definedName in workbook.Worksheets.Names)
            {
                // Consider only names that refer to a range (skip formulas, external refs)
                if (!string.IsNullOrEmpty(definedName.RefersTo) && definedName.RefersTo.StartsWith("="))
                {
                    // Obtain the range object for the name
                    AsposeRange range = definedName.GetRange();
                    if (range != null)
                    {
                        // Destination cell – use the first cell of the range
                        Cell destinationCell = range.Worksheet.Cells[range.FirstRow, range.FirstColumn];

                        // Create a bookmark entry for this named range
                        PdfBookmarkEntry entry = new PdfBookmarkEntry
                        {
                            Text = definedName.Text,   // Bookmark title = name of the range
                            Destination = destinationCell,
                            IsOpen = true
                        };

                        // Add to root's sub‑entries
                        rootBookmark.SubEntry.Add(entry);
                    }
                }
            }

            // ------------------------------------------------------------
            // Configure PDF save options with the constructed bookmark
            // ------------------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = rootBookmark
            };

            // Save the workbook as PDF (lifecycle rule: save)
            workbook.Save("NamedRangesBookmarks.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}