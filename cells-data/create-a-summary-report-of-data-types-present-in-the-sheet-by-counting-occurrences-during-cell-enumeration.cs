using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsDataTypeSummary
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate the worksheet with sample data of various types
            cells["A1"].PutValue(123);                     // Numeric
            cells["B1"].PutValue("Hello World");           // String
            cells["C1"].PutValue(DateTime.Now);            // DateTime
            cells["D1"].PutValue(true);                    // Boolean
            cells["E1"].PutValue(null);                    // Null (blank)
            cells["F1"].PutValue(45.67);                   // Numeric
            cells["A2"].PutValue("Another string");        // String
            cells["B2"].PutValue(false);                   // Boolean
            // Leave some cells empty to represent blanks

            // Dictionary to hold counts of each CellValueType
            Dictionary<CellValueType, int> typeCounts = new Dictionary<CellValueType, int>();

            // Initialize counts for all possible types
            foreach (CellValueType type in Enum.GetValues(typeof(CellValueType)))
            {
                typeCounts[type] = 0;
            }

            // Determine the used range of the worksheet
            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;

            // Enumerate cells within the used range and count value types
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = cells[row, col];
                    CellValueType cellType = cell.Type;
                    // Increment the count for the detected type
                    typeCounts[cellType]++;
                }
            }

            // Create a new worksheet to hold the summary report
            Worksheet summarySheet = workbook.Worksheets[workbook.Worksheets.Add()];
            summarySheet.Name = "DataTypeSummary";
            Cells summaryCells = summarySheet.Cells;

            // Write header
            summaryCells["A1"].PutValue("Cell Value Type");
            summaryCells["B1"].PutValue("Count");

            // Write the counts to the summary sheet
            int summaryRow = 1; // start from second row (index 1)
            foreach (var kvp in typeCounts)
            {
                // Only include types that actually appear (count > 0)
                if (kvp.Value > 0)
                {
                    summaryCells[summaryRow, 0].PutValue(kvp.Key.ToString());
                    summaryCells[summaryRow, 1].PutValue(kvp.Value);
                    summaryRow++;
                }
            }

            // Save the workbook to a file
            workbook.Save("DataTypeSummaryReport.xlsx");
        }
    }
}