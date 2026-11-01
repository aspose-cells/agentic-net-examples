// Title: Asynchronously refresh an Aspose.Cells PivotTable in a WinForms or WPF desktop app using C#
// Common Searches: C# Aspose.Cells refresh PivotTable on a background thread in WinForms | How to prevent UI freeze when refreshing an Excel PivotTable with Aspose.Cells | Async PivotTable.RefreshData example for WPF using Aspose.Cells | Load workbook and refresh first PivotTable without blocking the UI in C# | Task.Run Aspose.Cells PivotTable refresh in a desktop application
// Tags: asynchronous pivot table refresh Aspose.Cells | background thread RefreshData C# | winforms nonblocking pivot refresh | wpf excel workbook update async | task.run aspocells pivot refresh

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample loads an Excel workbook (or creates one if missing), accesses the first worksheet, checks for a PivotTable, and refreshes it on a background thread using Task.Run to call RefreshData. The asynchronous method prevents UI blocking in WinForms or WPF applications, and the workbook is saved after the refresh.
class Program
{
    // Entry point for a desktop application (e.g., WinForms/WPF)
    static async Task Main(string[] args)
    {
        try
        {
            // Verify that the input file exists before loading
            string inputPath = "input.xlsx";
            Workbook workbook;

            if (File.Exists(inputPath))
            {
                // Load the workbook (lifecycle rule: use provided load logic)
                workbook = new Workbook(inputPath);
            }
            else
            {
                // If the file is missing, create a new workbook with a default worksheet
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one PivotTable on the sheet
            if (sheet.PivotTables.Count > 0)
            {
                // Retrieve the first PivotTable
                PivotTable pivot = sheet.PivotTables[0];

                // Refresh the PivotTable without blocking the UI thread
                await RefreshPivotTableAsync(pivot);
            }
            else
            {
                Console.WriteLine("No PivotTable found on the first worksheet.");
            }

            // Save the workbook after the refresh (lifecycle rule: use provided save logic)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log or handle unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Asynchronous wrapper that runs the refresh on a background thread
    static Task RefreshPivotTableAsync(PivotTable pivotTable)
    {
        return Task.Run(() =>
        {
            // Refresh the underlying data source of the PivotTable
            pivotTable.RefreshData();

            // Optionally, refresh the data source if needed
            // pivotTable.RefreshDataSource();
        });
    }
}
