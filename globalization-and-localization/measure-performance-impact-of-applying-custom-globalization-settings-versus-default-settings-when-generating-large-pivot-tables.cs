using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

namespace AsposeCellsPerformanceDemo
{
    // Custom globalization settings for pivot tables
    public class CustomPivotGlobalizationSettings : PivotGlobalizationSettings
    {
        public override string GetTextOfTotal()
        {
            return "Custom Total";
        }

        public override string GetTextOfGrandTotal()
        {
            return "Custom Grand Total";
        }

        public override string GetTextOfDataFieldHeader()
        {
            return "Custom Data Header";
        }

        public override string GetTextOfProtectedName(string protectedName)
        {
            return protectedName + "_Custom";
        }
    }

    class Program
    {
        static void Main()
        {
            // Parameters for large dataset
            const int rowCount = 100_000; // number of data rows
            const int colCount = 5;       // number of data columns (including row header)

            // Measure default globalization performance
            Stopwatch swDefault = Stopwatch.StartNew();
            Workbook wbDefault = CreateWorkbookWithPivot(rowCount, colCount, useCustomSettings: false);
            swDefault.Stop();

            // Measure custom globalization performance
            Stopwatch swCustom = Stopwatch.StartNew();
            Workbook wbCustom = CreateWorkbookWithPivot(rowCount, colCount, useCustomSettings: true);
            swCustom.Stop();

            // Output results
            Console.WriteLine($"Default globalization time: {swDefault.ElapsedMilliseconds} ms");
            Console.WriteLine($"Custom globalization time: {swCustom.ElapsedMilliseconds} ms");

            // Save workbooks (optional, just to verify correctness)
            wbDefault.Save("Pivot_Default.xlsx");
            wbCustom.Save("Pivot_Custom.xlsx");
        }

        /// <summary>
        /// Creates a workbook, fills it with sample data, applies optional custom globalization,
        /// adds a pivot table, refreshes and calculates it.
        /// </summary>
        /// <param name="rows">Number of data rows to generate.</param>
        /// <param name="cols">Number of data columns (including the first column for row field).</param>
        /// <param name="useCustomSettings">If true, applies custom pivot globalization settings.</param>
        /// <returns>The populated workbook.</returns>
        private static Workbook CreateWorkbookWithPivot(int rows, int cols, bool useCustomSettings)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];
            Cells cells = dataSheet.Cells;

            // Populate header row
            for (int c = 0; c < cols; c++)
            {
                cells[0, c].PutValue($"Field{c + 1}");
            }

            // Populate large dataset
            Random rnd = new Random(0);
            for (int r = 1; r <= rows; r++)
            {
                // First column: categorical data (e.g., "CategoryA" to "CategoryZ")
                string category = "Category" + ((r - 1) % 26);
                cells[r, 0].PutValue(category);

                // Remaining columns: numeric data
                for (int c = 1; c < cols; c++)
                {
                    cells[r, c].PutValue(rnd.NextDouble() * 1000);
                }
            }

            // Apply custom globalization settings if requested
            if (useCustomSettings)
            {
                // Create an instance of custom pivot settings
                CustomPivotGlobalizationSettings pivotSettings = new CustomPivotGlobalizationSettings();

                // Assign to workbook's globalization settings
                workbook.Settings.GlobalizationSettings = new GlobalizationSettings();
                workbook.Settings.GlobalizationSettings.PivotSettings = pivotSettings;
            }

            // Define the data range for the pivot table
            string dataRange = $"A1:{CellIndexToName(rows, cols - 1)}";

            // Add a new worksheet for the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Add the pivot table
            int pivotIndex = pivotSheet.PivotTables.Add(dataRange, "A1", "LargePivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Add the first column as row field
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);

            // Add all numeric columns as data fields
            for (int c = 1; c < cols; c++)
            {
                pivotTable.AddFieldToArea(PivotFieldType.Data, c);
            }

            // Refresh and calculate the pivot table to materialize it
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            return workbook;
        }

        /// <summary>
        /// Converts zero‑based row and column indexes to an Excel cell name (e.g., (0,0) => "A1").
        /// Used to build the data range string.
        /// </summary>
        private static string CellIndexToName(int rowIndex, int columnIndex)
        {
            // Convert column index to letters
            string columnName = "";
            int dividend = columnIndex + 1;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            // Excel rows are 1‑based
            int excelRow = rowIndex + 1;
            return $"{columnName}{excelRow}";
        }
    }
}