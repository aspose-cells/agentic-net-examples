// Title: Asynchronously Refresh an Aspose.Cells PivotTable and Trigger a UI Callback in C#
// Description: Demonstrates how to create a workbook with a PivotTable, refresh its data on a background thread using Task.Run, recalculate the pivot, save the workbook, and then post a UI callback through a captured SynchronizationContext so the UI can be updated safely.
// Keywords: Aspose.Cells | PivotTable | RefreshData | CalculateData | async refresh | Task.Run | SynchronizationContext | C# | .NET | background thread | UI callback | WinForms | WPF | Excel automation
// Common Searches: how to refresh Aspose.Cells pivot table asynchronously | update UI after Aspose.Cells pivot refresh | use SynchronizationContext with Aspose.Cells | background thread pivot refresh C# | async PivotTable RefreshData Aspose.Cells
// Developer Intent: Refresh a PivotTable on a separate thread and notify the UI when the operation finishes.
// Use Cases: Refresh large PivotTables in a WinForms or WPF app without blocking the UI thread. | Run pivot data refresh in a Windows service or background worker and signal completion to a front‑end component. | Batch‑process many workbooks, refreshing each PivotTable asynchronously and persisting the results.
// AI Prompts: Generate C# code that asynchronously refreshes an Aspose.Cells PivotTable, saves the workbook, and invokes a supplied Action on the captured UI SynchronizationContext. | Create robust error‑handling for an async PivotTable refresh using Task.Run and ContinueWith with Aspose.Cells. | Show how to call RefreshPivotTableAsync from a WinForms button click, disabling the button during refresh and re‑enabling it in the UI callback.

using System;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshDemo
{
    // Simple UI simulation class
    // Demonstrates how to create a workbook with a PivotTable, refresh its data on a background thread using Task.Run, recalculate the pivot, save the workbook, and then post a UI callback through a captured SynchronizationContext so the UI can be updated safely.
    public class UiSimulator
    {
        // This method will be called after the pivot table refresh is complete
        public void OnRefreshCompleted()
        {
            Console.WriteLine("Pivot table refresh completed. UI can be updated now.");
            // Place UI update logic here (e.g., refresh a grid, enable buttons, etc.)
        }
    }

    public class Program
    {
        // Entry point
        public static void Main()
        {
            // Capture the synchronization context of the current thread (simulating UI thread)
            SynchronizationContext uiContext = SynchronizationContext.Current ?? new SynchronizationContext();

            // Create a workbook with sample data and a pivot table
            Workbook workbook = CreateWorkbookWithPivot();

            // Get the pivot table reference
            PivotTable pivotTable = workbook.Worksheets[0].PivotTables[0];

            // UI handler instance
            UiSimulator ui = new UiSimulator();

            // Start the refresh operation on a background thread
            RefreshPivotTableAsync(workbook, pivotTable, uiContext, ui.OnRefreshCompleted);

            // Prevent the console app from exiting immediately
            Console.WriteLine("Refresh started on background thread. Press any key to exit...");
            Console.ReadKey();
        }

        // Creates a workbook, fills data and adds a pivot table
        private static Workbook CreateWorkbookWithPivot()
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Sample data
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Sales");
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(120);
            cells["A3"].PutValue("Banana");
            cells["B3"].PutValue(150);
            cells["A4"].PutValue("Apple");
            cells["B4"].PutValue(80);
            cells["A5"].PutValue("Banana");
            cells["B5"].PutValue(70);

            // Add pivot table
            int pivotIndex = ws.PivotTables.Add("A1:B5", "D3", "SalesPivot");
            PivotTable pt = ws.PivotTables[pivotIndex];
            pt.AddFieldToArea(PivotFieldType.Row, 0);   // Product
            pt.AddFieldToArea(PivotFieldType.Data, 1);  // Sales

            // Initial calculation so the pivot shows data before refresh
            pt.CalculateData();

            // Save the initial workbook (optional, demonstrates lifecycle rule)
            try
            {
                wb.Save("InitialPivotWorkbook.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving initial workbook: {ex.Message}");
            }

            return wb;
        }

        // Performs pivot table refresh on a background thread and invokes a UI callback when done
        private static void RefreshPivotTableAsync(Workbook workbook, PivotTable pivotTable, SynchronizationContext uiContext, Action uiCallback)
        {
            Task.Run(() =>
            {
                try
                {
                    // Refresh the pivot table data
                    pivotTable.RefreshData();

                    // Recalculate the pivot table
                    pivotTable.CalculateData();

                    // Save the workbook after refresh (demonstrates lifecycle rule)
                    workbook.Save("RefreshedPivotWorkbook.xlsx");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error during pivot refresh: {ex.Message}");
                    throw; // Propagate to the continuation for fault handling
                }
            })
            .ContinueWith(t =>
            {
                // Handle any exception that occurred during the refresh task
                if (t.IsFaulted)
                {
                    Console.WriteLine("Refresh task failed.");
                    return;
                }

                // Post the callback to the captured UI synchronization context
                uiContext.Post(state => uiCallback?.Invoke(), null);
            });
        }
    }
}
