// Title: Refresh All Pivot Tables in Aspose.Cells (C#) After Updating Source Data
// Description: Demonstrates how to load or create a workbook, modify the source worksheet values, invoke Workbook.Worksheets.RefreshPivotTables() to recalculate every pivot table, and save the refreshed file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | RefreshPivotTables | C# | pivot table refresh | update source data | Excel workbook | programmatic pivot refresh | .NET | Workbook.Worksheets.RefreshPivotTables method | dynamic report generation
// Common Searches: Aspose.Cells refresh all pivot tables C# | How to update pivot tables after changing data in Aspose.Cells | RefreshPivotTables method example | Programmatically refresh Excel pivot tables .NET | Refresh pivot tables in workbook using Aspose.Cells
// Developer Intent: Recalculate every pivot table in a workbook after the underlying worksheet data has been changed programmatically.
// Use Cases: After adjusting sales figures in code, call RefreshPivotTables so summary sheets display the new totals before exporting. | Generate a dynamic financial report that modifies raw data and automatically updates all pivot analyses. | In an ETL workflow, alter data rows and ensure downstream pivot tables reflect the latest values by invoking RefreshPivotTables.
// AI Prompts: Write C# code with Aspose.Cells that updates cell values and refreshes all pivot tables in a workbook. | Show error‑handling patterns when calling Workbook.Worksheets.RefreshPivotTables on a file containing multiple pivots. | Compare Workbook.Worksheets.RefreshPivotTables with Worksheet.PivotTables[i].RefreshPivotTable in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to load or create a workbook, modify the source worksheet values, invoke Workbook.Worksheets.RefreshPivotTables() to recalculate every pivot table, and save the refreshed file using Aspose.Cells for .NET.
    public class RefreshAllPivotTablesDemo
    {
        public static void Run()
        {
            const string sourcePath = "source.xlsx";
            const string outputPath = "refreshed.xlsx";

            try
            {
                // Ensure source file exists; create a minimal workbook if missing
                if (!File.Exists(sourcePath))
                {
                    var wb = new Workbook();
                    var sourceDataSheet = wb.Worksheets[0];
                    sourceDataSheet.Name = "Data";

                    // Populate sample data
                    sourceDataSheet.Cells["A1"].PutValue("Category");
                    sourceDataSheet.Cells["B1"].PutValue("Amount");
                    sourceDataSheet.Cells["A2"].PutValue("A");
                    sourceDataSheet.Cells["B2"].PutValue(1000);
                    sourceDataSheet.Cells["A3"].PutValue("B");
                    sourceDataSheet.Cells["B3"].PutValue(2000);

                    // Add a simple pivot table for demonstration
                    var pivotSheet = wb.Worksheets.Add("Pivot");
                    int pivotIdx = pivotSheet.PivotTables.Add("=Data!A1:B3", "C3", "PivotTable1");
                    var pivotTable = pivotSheet.PivotTables[pivotIdx];

                    // Set Category as row field and Amount as data field
                    pivotTable.RowFields.Add(pivotTable.RowFields[0]);   // Category
                    pivotTable.DataFields.Add(pivotTable.DataFields[0]); // Amount

                    wb.Save(sourcePath);
                }

                // Load the workbook containing pivot tables
                var workbook = new Workbook(sourcePath);

                // Modify the underlying data source (example changes)
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Cells["B2"].PutValue(1500);
                dataSheet.Cells["B3"].PutValue(2500);

                // Refresh all pivot tables after data changes
                workbook.Worksheets.RefreshPivotTables();

                // Save the workbook with refreshed pivot tables
                workbook.Save(outputPath);
                Console.WriteLine($"Pivot tables refreshed and saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RefreshAllPivotTablesDemo.Run();
        }
    }
}
