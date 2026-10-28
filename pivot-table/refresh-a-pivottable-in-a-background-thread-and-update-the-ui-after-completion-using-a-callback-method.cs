// Title: Refresh an Aspose.Cells PivotTable on a background thread and update the UI via a callback in C#
// Common Searches: how to refresh an Aspose.Cells PivotTable without freezing the UI in a WinForms app | C# async method to refresh Excel pivot table using Aspose.Cells and invoke UI callback | run PivotTable.RefreshData on a background thread with Aspose.Cells in .NET | non‑blocking pivot table refresh using Aspose.Cells and Task.Run | update UI after asynchronous PivotTable refresh with Aspose.Cells
// Tags: asynchronous pivot table refresh Aspose.Cells | background thread workbook processing C# | PivotTable.RefreshData method Aspose.Cells | non‑blocking Excel pivot refresh .NET | UI callback after async workbook operation

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// Provides an async C# method that loads an Excel workbook with Aspose.Cells, locates a specific worksheet and PivotTable, calls RefreshData on a background thread via Task.Run, saves the file, and allows the calling code to update the UI through a callback after the await completes.
public class PivotTableRefresher
{
    // Refreshes the specified PivotTable asynchronously.
    public async Task RefreshPivotTableAsync(string workbookPath, string worksheetName, string pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(workbookPath))
            throw new ArgumentException("Workbook path is null or empty.", nameof(workbookPath));

        if (!File.Exists(workbookPath))
            throw new FileNotFoundException("Workbook file not found.", workbookPath);

        await Task.Run(() =>
        {
            try
            {
                // Load the workbook.
                var workbook = new Workbook(workbookPath);

                // Get the target worksheet.
                var sheet = workbook.Worksheets[worksheetName] ??
                            throw new ArgumentException($"Worksheet '{worksheetName}' not found.");

                // Get the target PivotTable.
                var pivotTable = sheet.PivotTables[pivotTableName] ??
                                 throw new ArgumentException($"PivotTable '{pivotTableName}' not found on worksheet '{worksheetName}'.");

                // Refresh the PivotTable data source.
                pivotTable.RefreshData();

                // Save the workbook.
                workbook.Save(workbookPath);
            }
            catch
            {
                // Propagate any exception to the calling async method.
                throw;
            }
        });
    }
}

public class Program
{
    // Example console usage.
    public static async Task Main(string[] args)
    {
        string filePath = @"C:\Data\Report.xlsx";
        string sheetName = "DataSheet";
        string pivotName = "PivotTable1";

        var refresher = new PivotTableRefresher();

        try
        {
            await refresher.RefreshPivotTableAsync(filePath, sheetName, pivotName);
            Console.WriteLine("PivotTable refresh completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
