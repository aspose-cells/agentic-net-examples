using System;
using Aspose.Cells;
using Aspose.Cells.Settings;

class Program
{
    static void Main()
    {
        // Create a new workbook with three worksheets
        Workbook workbook = new Workbook();
        workbook.Worksheets.Add(); // Sheet2
        workbook.Worksheets.Add(); // Sheet3

        // Apply custom globalization settings to change the total label for the Sum function
        SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();
        globalization.SetTotalName(ConsolidationFunction.Sum, "Localized Sum");
        workbook.Settings.GlobalizationSettings = globalization;

        // Populate each worksheet with sample data and add subtotals
        for (int wsIndex = 0; wsIndex < workbook.Worksheets.Count; wsIndex++)
        {
            Worksheet sheet = workbook.Worksheets[wsIndex];
            Cells cells = sheet.Cells;

            // Sample header
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Amount");

            // Sample rows (5 rows per sheet)
            object[,] data = new object[,]
            {
                { "Group A", 1200 },
                { "Group A", 800 },
                { "Group B", 1500 },
                { "Group B", 700 },
                { "Group C", 900 }
            };

            for (int r = 0; r < data.GetLength(0); r++)
            {
                cells[r + 1, 0].PutValue(data[r, 0]); // Category column
                cells[r + 1, 1].PutValue(data[r, 1]); // Amount column
            }

            // Define the range that contains the data (including header)
            int startRow = 0;
            int startCol = 0;
            int endRow = data.GetLength(0); // includes header row
            int endCol = 1;
            CellArea area = CellArea.CreateCellArea(startRow, startCol, endRow, endCol);

            // Add subtotals:
            // - Group by the first column (Category)
            // - Use SUM function on the second column (Amount)
            // - Replace existing subtotals, no page breaks, summary placed below data
            cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 1 }, true, false, true);
        }

        // Save the workbook
        workbook.Save("AllSheetsWithLocalizedSubtotals.xlsx");
    }
}