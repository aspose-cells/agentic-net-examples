// Title: Asynchronously refresh an Excel PivotTable in a .NET desktop app using Aspose.Cells
// Common Searches: how to refresh an Excel PivotTable without UI freeze using Aspose.Cells in C# | example of async pivot cache refresh with Aspose.Cells for WinForms | load workbook and refresh a named PivotTable on a background thread Aspose.Cells | C# Aspose.Cells refresh PivotTable data asynchronously and save workbook
// Tags: asynchronous pivot table refresh Aspose.Cells | background thread pivot cache update C# | non blocking UI workbook save Aspose.Cells | refresh named PivotTable .NET | Task.Run pivot table calculation Aspose.Cells

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

namespace PivotRefreshDemo
{
    // Provides a C# PivotHelper class with an async RefreshPivotTableAsync method that loads a workbook, validates the worksheet index, retrieves a PivotTable by name, refreshes its data and recalculates it on a background thread via Task.Run, then saves the file, handling missing files, invalid indexes, absent PivotTables, and save errors to keep the UI responsive.
    public class PivotHelper
    {
        // Asynchronously refreshes a PivotTable to keep the UI responsive.
        public async Task RefreshPivotTableAsync(string filePath, int worksheetIndex, string pivotTableName)
        {
            // Verify that the workbook file exists
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Workbook file not found: {filePath}");

            Workbook workbook;
            try
            {
                // Load the workbook
                workbook = new Workbook(filePath);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to load workbook.", ex);
            }

            // Validate worksheet index
            if (worksheetIndex < 0 || worksheetIndex >= workbook.Worksheets.Count)
                throw new ArgumentOutOfRangeException(nameof(worksheetIndex), "Worksheet index is out of range.");

            var worksheet = workbook.Worksheets[worksheetIndex];

            // Retrieve the PivotTable by its name
            var pivotTable = worksheet.PivotTables[pivotTableName];
            if (pivotTable == null)
                throw new ArgumentException($"PivotTable '{pivotTableName}' not found in worksheet index {worksheetIndex}.");

            // Refresh on a background thread to avoid UI blocking
            await Task.Run(() =>
            {
                // Refresh the pivot cache data
                pivotTable.RefreshData();

                // Recalculate the PivotTable after the refresh
                pivotTable.CalculateData();
            });

            try
            {
                // Save the workbook
                workbook.Save(filePath);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to save workbook.", ex);
            }
        }
    }

    class Program
    {
        // Entry point required for console application
        static async Task Main(string[] args)
        {
            // Example parameters (adjust as needed)
            string filePath = "sample.xlsx";
            int worksheetIndex = 0;
            string pivotTableName = "PivotTable1";

            var helper = new PivotHelper();

            try
            {
                await helper.RefreshPivotTableAsync(filePath, worksheetIndex, pivotTableName);
                Console.WriteLine("PivotTable refreshed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
