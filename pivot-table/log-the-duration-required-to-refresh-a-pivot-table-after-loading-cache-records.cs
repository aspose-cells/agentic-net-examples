using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshTiming
{
    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a workbook with sample data and a pivot table
            // -------------------------------------------------
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Amount";
            cells["A2"].Value = "Food";
            cells["B2"].Value = 120;
            cells["A3"].Value = "Food";
            cells["B3"].Value = 80;
            cells["A4"].Value = "Drink";
            cells["B4"].Value = 150;
            cells["A5"].Value = "Drink";
            cells["B5"].Value = 70;

            // Add a pivot table
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Amount

            // Initial refresh to build the cache
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook (cache is stored)
            string filePath = "PivotRefreshTiming.xlsx";
            wb.Save(filePath);

            // -------------------------------------------------
            // 2. Load the workbook and measure refresh duration
            // -------------------------------------------------
            Workbook loadedWb = new Workbook(filePath);
            Worksheet loadedSheet = loadedWb.Worksheets[0];
            PivotTable loadedPivot = loadedSheet.PivotTables[0];

            // Start timing
            Stopwatch sw = Stopwatch.StartNew();

            // Refresh the pivot table (this will rebuild the cache from source data)
            loadedPivot.RefreshData();

            // Stop timing
            sw.Stop();

            // Output the duration
            Console.WriteLine($"Pivot table refresh duration: {sw.ElapsedMilliseconds} ms");

            // Optional: calculate data so the refreshed values appear in the sheet
            loadedPivot.CalculateData();

            // Save the workbook after refresh (if needed)
            loadedWb.Save("PivotRefreshTiming_Refreshed.xlsx");
        }
    }
}