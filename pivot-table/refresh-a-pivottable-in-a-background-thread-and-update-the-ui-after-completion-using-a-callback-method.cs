using System;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static async Task Main()
    {
        // -------------------- Create workbook and sample data --------------------
        Workbook workbook = new Workbook();                         // create workbook
        Worksheet sheet = workbook.Worksheets[0];

        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(200);
        sheet.Cells["A4"].PutValue("Apple");
        sheet.Cells["B4"].PutValue(150);

        // -------------------- Add a pivot table --------------------
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product column
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales column

        // Initial calculation so the pivot has data before refresh
        pivotTable.CalculateData();

        // -------------------- Define UI callback --------------------
        // In a real UI (WinForms/WPF) this would update controls.
        // Here we simply write to the console and save the workbook.
        Action refreshCompletedCallback = () =>
        {
            Console.WriteLine("Pivot table refresh completed.");
            workbook.Save("PivotRefreshed.xlsx");   // save after refresh
        };

        // -------------------- Refresh pivot in background thread --------------------
        await RefreshPivotAsync(pivotTable, refreshCompletedCallback);

        // Keep console open to view output
        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }

    // Runs RefreshData and CalculateData on a background thread,
    // then invokes the supplied callback on the original (UI) thread.
    static async Task RefreshPivotAsync(PivotTable pivot, Action callback)
    {
        await Task.Run(() =>
        {
            // Refresh data source of the pivot table
            pivot.RefreshData();

            // Recalculate the pivot table after data refresh
            pivot.CalculateData();
        });

        // Invoke UI callback after background work is done
        callback?.Invoke();
    }
}