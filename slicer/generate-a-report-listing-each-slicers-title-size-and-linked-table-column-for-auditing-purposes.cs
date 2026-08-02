using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;
using Aspose.Cells.Pivot;

namespace SlicerAuditReport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -----------------------------
            // Sample data and slicer setup
            // -----------------------------
            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("A");
            sheet.Cells["B4"].PutValue(30);

            // Add a table based on the data
            int tableIdx = sheet.ListObjects.Add(0, 0, 3, 1, true);
            ListObject table = sheet.ListObjects[tableIdx];

            // Add a slicer linked to the first column of the table
            int slicerIdx = sheet.Slicers.Add(table, 0, "D2");
            Slicer slicer = sheet.Slicers[slicerIdx];
            slicer.Caption = "Category Slicer";

            // ---------------------------------
            // Create a worksheet for the audit report
            // ---------------------------------
            Worksheet reportSheet = workbook.Worksheets.Add("SlicerReport");

            // Write header row
            reportSheet.Cells["A1"].PutValue("Worksheet");
            reportSheet.Cells["B1"].PutValue("Slicer Name");
            reportSheet.Cells["C1"].PutValue("Caption");
            reportSheet.Cells["D1"].PutValue("Width (px)");
            reportSheet.Cells["E1"].PutValue("Height (px)");
            reportSheet.Cells["F1"].PutValue("Source Name");

            int reportRow = 1; // zero‑based index; row 1 is the second row (after headers)

            // Iterate through all worksheets (except the report sheet itself)
            for (int wsIdx = 0; wsIdx < workbook.Worksheets.Count; wsIdx++)
            {
                Worksheet ws = workbook.Worksheets[wsIdx];
                if (ws.Name == reportSheet.Name) continue; // skip the report sheet

                SlicerCollection slicers = ws.Slicers;
                for (int i = 0; i < slicers.Count; i++)
                {
                    Slicer s = slicers[i];

                    // Gather required information
                    string worksheetName = ws.Name;
                    string slicerName = s.Name;
                    string caption = s.Caption;
                    int widthPx = s.WidthPixel;
                    int heightPx = s.HeightPixel;
                    string sourceName = s.SlicerCache?.SourceName ?? "N/A";

                    // Write data to the report sheet
                    reportSheet.Cells[reportRow, 0].PutValue(worksheetName);
                    reportSheet.Cells[reportRow, 1].PutValue(slicerName);
                    reportSheet.Cells[reportRow, 2].PutValue(caption);
                    reportSheet.Cells[reportRow, 3].PutValue(widthPx);
                    reportSheet.Cells[reportRow, 4].PutValue(heightPx);
                    reportSheet.Cells[reportRow, 5].PutValue(sourceName);

                    reportRow++;
                }
            }

            // Save the workbook with the audit report
            workbook.Save("SlicerAuditReport.xlsx");
        }
    }
}