using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsTimelineFromTSV
{
    class Program
    {
        static void Main()
        {
            // Path to the TSV file containing events and timestamps
            string tsvPath = "events.tsv"; // Example: Event<TAB>Timestamp per line
            // Output Excel file
            string outputPath = "TimelineFromTSV.xlsx";

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Load TSV data into the worksheet starting at cell A1
            if (!File.Exists(tsvPath))
            {
                Console.WriteLine($"TSV file not found: {tsvPath}");
                return;
            }

            string[] lines = File.ReadAllLines(tsvPath);
            for (int row = 0; row < lines.Length; row++)
            {
                // Split each line by tab character
                string[] parts = lines[row].Split('\t');
                for (int col = 0; col < parts.Length; col++)
                {
                    // Try to parse as DateTime; otherwise store as string
                    if (DateTime.TryParse(parts[col], out DateTime dt))
                        cells[row, col].PutValue(dt);
                    else
                        cells[row, col].PutValue(parts[col]);
                }
            }

            // Assume first column is "Event", second column is "Timestamp"
            // Create a PivotTable using the imported range
            // Determine the used range
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;
            string sourceRange = $"A1:{CellIndexToName(maxRow, maxCol)}";

            // Place the PivotTable at cell D1
            int pivotIdx = sheet.PivotTables.Add(sourceRange, "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];

            // Add the Timestamp column as a Row field (required for Timeline)
            pivot.AddFieldToArea(PivotFieldType.Row, "Timestamp");
            // Optionally add Event as a Data field to show counts
            pivot.AddFieldToArea(PivotFieldType.Data, "Event");

            // Refresh the PivotTable data
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a Timeline linked to the PivotTable, positioned at cell G1, using the Timestamp field
            int timelineIdx = sheet.Timelines.Add(pivot, "G1", "Timestamp");
            // Optionally set a name for the Timeline
            sheet.Timelines[timelineIdx].Name = "EventTimeline";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }

        // Helper to convert zero‑based row/column indexes to Excel cell name (e.g., 0,0 -> A1)
        static string CellIndexToName(int rowIndex, int columnIndex)
        {
            int dividend = columnIndex + 1;
            string columnName = string.Empty;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return $"{columnName}{rowIndex + 1}";
        }
    }
}