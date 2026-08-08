// Title: C# – Aspose.Cells – Create Slicer Audit Report (title, size, linked column)
// Description: Loads an existing workbook, adds a "SlicerReport" worksheet, writes headers, then scans every worksheet (except the report) to capture each slicer's name, title, width (px), height (px) and the column it is linked to via the slicer cache. The collected data is written row‑by‑row and the workbook is saved as "SlicerAuditReport.xlsx".
// Keywords: Aspose.Cells | C# | slicer audit | slicer title extraction | slicer dimensions | linked column retrieval | Excel automation | listobject slicer cache | dashboard validation | generate slicer report
// Common Searches: Aspose.Cells list all slicers in a workbook | C# get slicer width and height using Aspose.Cells | retrieve slicer linked column Aspose.Cells .NET | create Excel report of slicer metadata C# | how to audit slicer settings with Aspose.Cells
// Developer Intent: Produce an Excel file that enumerates every slicer’s worksheet, name, title, pixel dimensions, and associated table column.
// Use Cases: Validate slicer titles and sizes across a multi‑sheet dashboard for UI consistency. | Document which table columns are bound to slicers to ensure data integrity before publishing. | Generate a quick inventory of slicer configurations for compliance or audit trails.
// AI Prompts: Write C# code with Aspose.Cells that extracts slicer title, width, height, and linked column into a new worksheet. | Explain safe methods to obtain the linked column name from a slicer’s cache and how to handle missing information. | Suggest additional columns for the slicer audit report, such as slicer style, position, or current filter state.

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;
using Aspose.Cells.Tables;

// Loads an existing workbook, adds a "SlicerReport" worksheet, writes headers, then scans every worksheet (except the report) to capture each slicer's name, title, width (px), height (px) and the column it is linked to via the slicer cache. The collected data is written row‑by‑row and the workbook is saved as "SlicerAuditReport.xlsx".
class SlicerAuditReport
{
    static void Main()
    {
        // Load the workbook that contains slicers
        Workbook workbook = new Workbook("input.xlsx");

        // Add a new worksheet for the report
        int reportIndex = workbook.Worksheets.Add();
        Worksheet reportSheet = workbook.Worksheets[reportIndex];
        reportSheet.Name = "SlicerReport";

        // Write header row
        Cells reportCells = reportSheet.Cells;
        reportCells[0, 0].PutValue("Worksheet");
        reportCells[0, 1].PutValue("Slicer Name");
        reportCells[0, 2].PutValue("Title");
        reportCells[0, 3].PutValue("Width (px)");
        reportCells[0, 4].PutValue("Height (px)");
        reportCells[0, 5].PutValue("Linked Column");

        int currentRow = 1;

        // Iterate through all worksheets and their slicers
        foreach (Worksheet ws in workbook.Worksheets)
        {
            // Skip the report sheet itself
            if (ws.Name == reportSheet.Name) continue;

            SlicerCollection slicers = ws.Slicers;
            for (int i = 0; i < slicers.Count; i++)
            {
                Slicer slicer = slicers[i];

                // Basic slicer information
                string worksheetName = ws.Name;
                string slicerName = slicer.Name;
                string title = slicer.Title;               // Title (may be empty if not set)
                int widthPx = slicer.WidthPixel;
                int heightPx = slicer.HeightPixel;

                // Attempt to retrieve the linked column name via the slicer cache
                string linkedColumn = string.Empty;
                try
                {
                    // For slicers based on a ListObject, SourceName usually contains the column name
                    linkedColumn = slicer.SlicerCache.SourceName;
                }
                catch
                {
                    // If unavailable, leave the field empty
                }

                // Write the data to the report sheet
                reportCells[currentRow, 0].PutValue(worksheetName);
                reportCells[currentRow, 1].PutValue(slicerName);
                reportCells[currentRow, 2].PutValue(title);
                reportCells[currentRow, 3].PutValue(widthPx);
                reportCells[currentRow, 4].PutValue(heightPx);
                reportCells[currentRow, 5].PutValue(linkedColumn);

                currentRow++;
            }
        }

        // Save the workbook with the new report
        workbook.Save("SlicerAuditReport.xlsx");
    }
}
