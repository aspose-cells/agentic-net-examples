using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

class SlicerAuditReport
{
    static void Main()
    {
        // Load the workbook that contains slicers
        Workbook workbook = new Workbook("input.xlsx");

        // Add a new worksheet to hold the audit report
        int reportSheetIndex = workbook.Worksheets.Add();
        Worksheet reportSheet = workbook.Worksheets[reportSheetIndex];
        reportSheet.Name = "SlicerReport";

        // Write header row
        reportSheet.Cells["A1"].PutValue("Worksheet");
        reportSheet.Cells["B1"].PutValue("Slicer Name");
        reportSheet.Cells["C1"].PutValue("Caption");
        reportSheet.Cells["D1"].PutValue("Width (px)");
        reportSheet.Cells["E1"].PutValue("Height (px)");
        reportSheet.Cells["F1"].PutValue("Linked Source");

        int currentRow = 1; // zero‑based index for the next data row

        // Iterate through all worksheets (except the report sheet itself)
        foreach (Worksheet ws in workbook.Worksheets)
        {
            if (ws.Name == "SlicerReport")
                continue;

            SlicerCollection slicers = ws.Slicers;

            // Process each slicer on the worksheet
            for (int i = 0; i < slicers.Count; i++)
            {
                Slicer slicer = slicers[i];

                // Worksheet name
                reportSheet.Cells[currentRow, 0].PutValue(ws.Name);
                // Slicer object name
                reportSheet.Cells[currentRow, 1].PutValue(slicer.Name);
                // Caption (used as the title)
                reportSheet.Cells[currentRow, 2].PutValue(slicer.Caption);
                // Size in pixels
                reportSheet.Cells[currentRow, 3].PutValue(slicer.WidthPixel);
                reportSheet.Cells[currentRow, 4].PutValue(slicer.HeightPixel);
                // Linked source (table or pivot field name)
                string linkedSource = slicer.SlicerCache != null ? slicer.SlicerCache.SourceName : string.Empty;
                reportSheet.Cells[currentRow, 5].PutValue(linkedSource);

                currentRow++;
            }
        }

        // Save the workbook with the added report sheet
        workbook.Save("SlicerReport.xlsx");
    }
}