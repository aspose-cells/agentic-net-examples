using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsDiffExample
{
    class Program
    {
        static void Main()
        {
            // Paths to the two workbooks to compare
            string workbookPath1 = "Workbook1.xlsx";
            string workbookPath2 = "Workbook2.xlsx";

            // Load the workbooks (create rule)
            Workbook wb1 = new Workbook(workbookPath1);
            Workbook wb2 = new Workbook(workbookPath2);

            // Assume we compare the first worksheet of each workbook
            Worksheet ws1 = wb1.Worksheets[0];
            Worksheet ws2 = wb2.Worksheets[0];

            // Add a new worksheet to hold the diff report (create rule)
            int diffSheetIndex = wb1.Worksheets.Add();
            Worksheet diffWs = wb1.Worksheets[diffSheetIndex];
            diffWs.Name = "DiffReport";

            // Write header row in the diff sheet
            diffWs.Cells["A1"].PutValue("Cell Address");
            diffWs.Cells["B1"].PutValue("Workbook1 Value");
            diffWs.Cells["C1"].PutValue("Workbook2 Value");

            int reportRow = 1; // zero‑based index; start after header

            // Build a lookup for cells in the second worksheet using its enumerator
            // Key: cell name (e.g., "A1"), Value: Cell object
            var ws2CellMap = new System.Collections.Generic.Dictionary<string, Cell>(StringComparer.OrdinalIgnoreCase);
            IEnumerator ws2Enum = ws2.Cells.GetEnumerator(); // Cells.GetEnumerator rule
            while (ws2Enum.MoveNext())
            {
                Cell cell = (Cell)ws2Enum.Current;
                if (cell != null && !string.IsNullOrEmpty(cell.Name))
                {
                    ws2CellMap[cell.Name] = cell;
                }
            }

            // Enumerate cells of the first worksheet and compare with the second
            IEnumerator ws1Enum = ws1.Cells.GetEnumerator(); // Cells.GetEnumerator rule
            while (ws1Enum.MoveNext())
            {
                Cell cell1 = (Cell)ws1Enum.Current;
                if (cell1 == null || string.IsNullOrEmpty(cell1.Name))
                    continue;

                // Try to find the corresponding cell in worksheet 2
                ws2CellMap.TryGetValue(cell1.Name, out Cell cell2);

                object val1 = cell1.Value;
                object val2 = cell2?.Value;

                // Determine if values are different (handle nulls)
                bool areDifferent = false;
                if (val1 == null && val2 == null)
                {
                    areDifferent = false;
                }
                else if (val1 == null || val2 == null)
                {
                    areDifferent = true;
                }
                else
                {
                    areDifferent = !val1.Equals(val2);
                }

                if (areDifferent)
                {
                    // Record the mismatch in the diff worksheet
                    diffWs.Cells[reportRow, 0].PutValue(cell1.Name);
                    diffWs.Cells[reportRow, 1].PutValue(val1);
                    diffWs.Cells[reportRow, 2].PutValue(val2);
                    reportRow++;
                }

                // Remove the entry from the map so that remaining entries represent cells only in ws2
                if (cell2 != null)
                {
                    ws2CellMap.Remove(cell1.Name);
                }
            }

            // Any remaining cells in ws2CellMap exist only in workbook2
            foreach (var kvp in ws2CellMap)
            {
                Cell cell2 = kvp.Value;
                diffWs.Cells[reportRow, 0].PutValue(cell2.Name);
                diffWs.Cells[reportRow, 1].PutValue(null); // No value in workbook1
                diffWs.Cells[reportRow, 2].PutValue(cell2.Value);
                reportRow++;
            }

            // Auto‑fit columns for better readability
            diffWs.AutoFitColumns();

            // Save the workbook containing the diff report (save rule)
            wb1.Save("DiffReportOutput.xlsx");
        }
    }
}