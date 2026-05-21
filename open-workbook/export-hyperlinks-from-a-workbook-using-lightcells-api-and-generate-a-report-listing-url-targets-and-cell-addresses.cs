using System;
using Aspose.Cells;
using Aspose.Cells.Utility; // For CellsHelper

class HyperlinkExport
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook sourceWorkbook = new Workbook("input.xlsx");

        // Create a new workbook for the report
        Workbook reportWorkbook = new Workbook();
        Worksheet reportSheet = reportWorkbook.Worksheets[0];
        reportSheet.Name = "HyperlinkReport";

        // Write header row
        reportSheet.Cells["A1"].PutValue("Worksheet");
        reportSheet.Cells["B1"].PutValue("Cell");
        reportSheet.Cells["C1"].PutValue("URL");

        int reportRow = 1; // zero‑based index; row 1 is the second row (after header)

        // Iterate through all worksheets in the source workbook
        foreach (Worksheet ws in sourceWorkbook.Worksheets)
        {
            // Access the HyperlinkCollection of the current worksheet
            HyperlinkCollection links = ws.Hyperlinks;

            // Enumerate each hyperlink
            foreach (Hyperlink link in links)
            {
                // Get the address (URL or internal reference)
                string url = link.Address;

                // Determine the top‑left cell of the hyperlink range
                int startRow = link.Area.StartRow;
                int startColumn = link.Area.StartColumn;

                // Convert row/column indices to the A1 style cell name
                string cellName = CellsHelper.CellIndexToName(startRow, startColumn);

                // Write data to the report sheet
                reportSheet.Cells[reportRow, 0].PutValue(ws.Name);   // Worksheet name
                reportSheet.Cells[reportRow, 1].PutValue(cellName); // Cell address
                reportSheet.Cells[reportRow, 2].PutValue(url);      // Hyperlink URL

                reportRow++;
            }
        }

        // Save the report workbook
        reportWorkbook.Save("HyperlinkReport.xlsx");
    }
}