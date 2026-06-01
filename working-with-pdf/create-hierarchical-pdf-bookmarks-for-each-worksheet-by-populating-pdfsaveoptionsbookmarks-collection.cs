using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfBookmarks
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add several worksheets
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Name = "Summary";
            Worksheet sheet1 = workbook.Worksheets.Add("Sales");
            Worksheet sheet2 = workbook.Worksheets.Add("Inventory");
            Worksheet sheet3 = workbook.Worksheets.Add("Employees");

            // Put a title in each sheet – this cell will be the bookmark destination
            workbook.Worksheets[0].Cells["A1"].PutValue("Summary Overview");
            sheet1.Cells["A1"].PutValue("Sales Data");
            sheet2.Cells["A1"].PutValue("Inventory List");
            sheet3.Cells["A1"].PutValue("Employee Records");

            // Create the root bookmark entry for the whole workbook
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = "Workbook",
                Destination = workbook.Worksheets[0].Cells["A1"], // point to first sheet
                IsOpen = true,
                SubEntry = new ArrayList()
            };

            // Helper method to create a bookmark entry for a worksheet
            PdfBookmarkEntry CreateSheetBookmark(Worksheet ws)
            {
                return new PdfBookmarkEntry
                {
                    Text = ws.Name,
                    Destination = ws.Cells["A1"],
                    // SubEntry can be populated further if deeper hierarchy is needed
                };
            }

            // Add a bookmark for each worksheet as a child of the root
            rootBookmark.SubEntry.Add(CreateSheetBookmark(workbook.Worksheets[0])); // Summary
            rootBookmark.SubEntry.Add(CreateSheetBookmark(sheet1));
            rootBookmark.SubEntry.Add(CreateSheetBookmark(sheet2));
            rootBookmark.SubEntry.Add(CreateSheetBookmark(sheet3));

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Export the document structure so that bookmarks are retained
                ExportDocumentStructure = true,
                // Assign the hierarchical bookmark tree
                Bookmark = rootBookmark
            };

            // Save the workbook as PDF with the defined bookmarks
            workbook.Save("WorkbookWithBookmarks.pdf", pdfOptions);
        }
    }
}