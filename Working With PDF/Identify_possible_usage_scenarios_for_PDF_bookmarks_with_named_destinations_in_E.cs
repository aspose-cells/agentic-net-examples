using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfBookmarksDemo
{
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Prepare a workbook with multiple worksheets and sample data
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();

            // Worksheet 1 – Overview
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Overview";
            sheet1.Cells["A1"].PutValue("Overview Section");
            sheet1.Cells["A2"].PutValue("This is the first page of the PDF.");

            // Worksheet 2 – Details
            Worksheet sheet2 = workbook.Worksheets.Add("Details");
            sheet2.Cells["A1"].PutValue("Details Section");
            sheet2.Cells["A2"].PutValue("Detailed information goes here.");

            // Worksheet 3 – Summary
            Worksheet sheet3 = workbook.Worksheets.Add("Summary");
            sheet3.Cells["A1"].PutValue("Summary Section");
            sheet3.Cells["A2"].PutValue("Summary of the document.");

            // ------------------------------------------------------------
            // 2. Scenario A – Simple root bookmark pointing to a cell
            // ------------------------------------------------------------
            PdfBookmarkEntry simpleRoot = new PdfBookmarkEntry
            {
                Text = "Overview",
                Destination = sheet1.Cells["A1"], // Direct link to cell A1 on Overview sheet
                IsOpen = true                     // Expanded when PDF is opened
            };

            // ------------------------------------------------------------
            // 3. Scenario B – Hierarchical bookmarks (root + sub‑bookmarks)
            // ------------------------------------------------------------
            PdfBookmarkEntry hierarchicalRoot = new PdfBookmarkEntry
            {
                Text = "Document Sections",
                Destination = sheet1.Cells["A1"],
                IsOpen = true,
                SubEntry = new ArrayList()
            };

            PdfBookmarkEntry subDetails = new PdfBookmarkEntry
            {
                Text = "Details",
                Destination = sheet2.Cells["A1"]
            };

            PdfBookmarkEntry subSummary = new PdfBookmarkEntry
            {
                Text = "Summary",
                Destination = sheet3.Cells["A1"]
            };

            // Add sub‑bookmarks to the root entry
            hierarchicalRoot.SubEntry.Add(subDetails);
            hierarchicalRoot.SubEntry.Add(subSummary);

            // ------------------------------------------------------------
            // 4. Scenario C – Named destinations using DestinationName
            //    (Useful when external PDF links need to reference a specific location)
            // ------------------------------------------------------------
            PdfBookmarkEntry namedDestinationRoot = new PdfBookmarkEntry
            {
                Text = "Named Destinations",
                Destination = sheet1.Cells["A1"],
                IsOpen = true,
                SubEntry = new ArrayList()
            };

            PdfBookmarkEntry namedDest1 = new PdfBookmarkEntry
            {
                Text = "Primary Destination",
                Destination = sheet2.Cells["A1"],
                DestinationName = "DetailsSection" // This creates a named destination
            };

            PdfBookmarkEntry namedDest2 = new PdfBookmarkEntry
            {
                Text = "Secondary Destination",
                Destination = sheet3.Cells["A1"],
                DestinationName = "SummarySection"
            };

            namedDestinationRoot.SubEntry.Add(namedDest1);
            namedDestinationRoot.SubEntry.Add(namedDest2);

            // ------------------------------------------------------------
            // 5. Scenario D – Hidden root bookmark (Text = null)
            //    Children are inserted at the top level, useful for flattening hierarchy.
            // ------------------------------------------------------------
            PdfBookmarkEntry hiddenRoot = new PdfBookmarkEntry
            {
                Text = null, // Hidden entry
                SubEntry = new ArrayList()
            };

            PdfBookmarkEntry flatBookmark1 = new PdfBookmarkEntry
            {
                Text = "Flat Overview",
                Destination = sheet1.Cells["A1"]
            };

            PdfBookmarkEntry flatBookmark2 = new PdfBookmarkEntry
            {
                Text = "Flat Details",
                Destination = sheet2.Cells["A1"]
            };

            hiddenRoot.SubEntry.Add(flatBookmark1);
            hiddenRoot.SubEntry.Add(flatBookmark2);

            // ------------------------------------------------------------
            // 6. Scenario E – Collapsed bookmark (IsOpen = false)
            //    The bookmark appears collapsed in the PDF bookmark pane.
            // ------------------------------------------------------------
            PdfBookmarkEntry collapsedRoot = new PdfBookmarkEntry
            {
                Text = "Collapsed Section",
                Destination = sheet3.Cells["A1"],
                IsOpen = false // Collapsed by default
            };

            // ------------------------------------------------------------
            // 7. Combine all scenarios into a single bookmark tree
            //    The root entry can be hidden (Text = null) to merge all top‑level items.
            // ------------------------------------------------------------
            PdfBookmarkEntry masterRoot = new PdfBookmarkEntry
            {
                Text = null, // Hidden root – all children appear at top level
                SubEntry = new ArrayList()
            };

            masterRoot.SubEntry.Add(simpleRoot);
            masterRoot.SubEntry.Add(hierarchicalRoot);
            masterRoot.SubEntry.Add(namedDestinationRoot);
            masterRoot.SubEntry.Add(hiddenRoot);
            masterRoot.SubEntry.Add(collapsedRoot);

            // ------------------------------------------------------------
            // 8. Configure PDF save options
            //    - Assign the bookmark tree
            //    - Export document structure (helps PDF readers build the outline)
            // ------------------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = masterRoot,
                ExportDocumentStructure = true
            };

            // ------------------------------------------------------------
            // 9. Save the workbook as PDF with the defined bookmarks
            // ------------------------------------------------------------
            workbook.Save("ExcelToPdf_WithBookmarks.pdf", pdfOptions);

            Console.WriteLine("PDF generated with various bookmark scenarios.");
        }
    }
}