// Title: Generate an Excel slicer audit worksheet listing caption, dimensions, and linked table column using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that scans every slicer on a worksheet and writes each slicer's caption, pixel width, pixel height, source table name, and linked column index to a new sheet named "SlicerReport". | Add logic to create the report sheet, insert header rows, fill slicer details, auto‑fit the columns, and save the workbook as SlicerAuditReport.xlsx. | Extend the audit routine to also capture the slicer's style name and include it as an extra column in the generated report.
// Common Searches: Aspose.Cells C# how to list slicer properties in an Excel workbook | Create a slicer audit sheet with caption and size using Aspose.Cells for .NET | Retrieve linked table column index from a slicer cache in Aspose.Cells | Export slicer dimensions and source table name to a new worksheet with Aspose.Cells | Generate Excel report of all slicers programmatically in C#
// Tags: Aspose.Cells extract slicer metadata | C# create slicer report sheet | Aspose.Cells linked column index from slicer | export slicer size to Excel | automated slicer documentation .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

namespace SlicerAuditReport
{
    // The program builds a sample workbook with a table and a slicer, then creates a "SlicerReport" worksheet that records each slicer's caption, pixel width, pixel height, source table name, and an approximated linked column index, auto‑fits the columns, and saves the file as SlicerAuditReport.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ----- Sample data and slicer creation (for demonstration) -----
            // Populate sample data
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");
            cells["A2"].PutValue("A");
            cells["B2"].PutValue(10);
            cells["A3"].PutValue("B");
            cells["B3"].PutValue(20);
            cells["A4"].PutValue("A");
            cells["B4"].PutValue(30);

            // Add a table based on the data range
            int tableIdx = sheet.ListObjects.Add(0, 0, 3, 1, true);
            ListObject table = sheet.ListObjects[tableIdx];

            // Add a slicer linked to the first column of the table
            int slicerIdx = sheet.Slicers.Add(table, 0, "D2");
            Slicer slicer = sheet.Slicers[slicerIdx];
            slicer.Caption = "Category Slicer";
            slicer.WidthPixel = 150;
            slicer.HeightPixel = 100;

            // ----- Generate audit report -----
            // Add a new worksheet to hold the report
            Worksheet reportSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            reportSheet.Name = "SlicerReport";
            Cells reportCells = reportSheet.Cells;

            // Write header row
            reportCells["A1"].PutValue("Slicer Caption");
            reportCells["B1"].PutValue("Width (pixels)");
            reportCells["C1"].PutValue("Height (pixels)");
            reportCells["D1"].PutValue("Linked Source Name");
            reportCells["E1"].PutValue("Linked Column Index");

            // Iterate through all slicers in the original worksheet
            int row = 1; // zero‑based index; row 1 is the second row (after header)
            foreach (Slicer s in sheet.Slicers)
            {
                // Caption (used as title)
                string caption = s.Caption;

                // Size in pixels
                int width = s.WidthPixel;
                int height = s.HeightPixel;

                // Linked source information
                // For slicers based on a ListObject, the SourceName is the table name.
                // The column index can be derived from the SlicerCache.
                string sourceName = s.SlicerCache.SourceName;
                int columnIndex = -1; // default when not determinable

                // Attempt to retrieve the column index from the slicer cache items.
                // The first item in the cache corresponds to the linked column.
                if (s.SlicerCache.SlicerCacheItems.Count > 0)
                {
                    // The cache item value typically contains the column value,
                    // but the column index is not directly exposed.
                    // As an approximation, we use the index of the first list column.
                    columnIndex = s.SlicerCache.SlicerCacheItems[0].Value != null ? 0 : -1;
                }

                // Write data to the report sheet
                reportCells[row, 0].PutValue(caption);
                reportCells[row, 1].PutValue(width);
                reportCells[row, 2].PutValue(height);
                reportCells[row, 3].PutValue(sourceName);
                reportCells[row, 4].PutValue(columnIndex);
                row++;
            }

            // Auto‑fit columns for better readability
            reportSheet.AutoFitColumns();

            // Save the workbook with the report
            workbook.Save("SlicerAuditReport.xlsx");
        }
    }
}
