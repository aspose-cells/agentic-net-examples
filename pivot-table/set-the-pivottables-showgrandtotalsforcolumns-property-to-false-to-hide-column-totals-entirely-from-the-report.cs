// Title: C# example: hide column grand totals in an Aspose.Cells PivotTable
// AI Prompts: Write C# code that creates a workbook, adds sample data, builds a PivotTable, and suppresses the column summary values by setting the appropriate flag to false. | Show how to refresh and recalculate a PivotTable after the column summary values have been suppressed, then save the workbook to a given path. | Provide a complete C# snippet that configures a PivotTable, hides column summary values, and writes the resulting file to disk without errors.
// Common Searches: Aspose.Cells C# hide column grand totals in pivot table | Set ShowColumnGrandTotals to false in Aspose.Cells PivotTable | Remove column totals from Excel pivot using Aspose.Cells .NET | How to disable column grand totals when creating a PivotTable with Aspose.Cells | C# Aspose.Cells example for hiding column totals in pivot report
// Tags: Aspose.Cells PivotTable total visibility control | Aspose.Cells column grand total flag usage | Excel pivot column summary suppression with Aspose | PivotTable RefreshData after total visibility change | Aspose.Cells workbook save with customized pivot totals

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The sample creates a workbook, populates it with sales data, adds a PivotTable, disables column grand totals via the relevant flag, refreshes and calculates the pivot, and saves the result as HideColumnGrandTotalsDemo.xlsx.
    public class HideColumnGrandTotalsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                Cells cells = sheet.Cells;
                cells["A1"].Value = "Category";
                cells["B1"].Value = "Region";
                cells["C1"].Value = "Sales";

                cells["A2"].Value = "Electronics";
                cells["B2"].Value = "North";
                cells["C2"].Value = 1200;

                cells["A3"].Value = "Electronics";
                cells["B3"].Value = "South";
                cells["C3"].Value = 1500;

                cells["A4"].Value = "Clothing";
                cells["B4"].Value = "North";
                cells["C4"].Value = 800;

                cells["A5"].Value = "Clothing";
                cells["B5"].Value = "South";
                cells["C5"].Value = 950;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table: rows = Category, columns = Region, data = Sales
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Category as row field
                pivotTable.AddFieldToArea(PivotFieldType.Column, 1);   // Region as column field
                pivotTable.AddFieldToArea(PivotFieldType.Data, 2);    // Sales as data field

                // Hide column grand totals
                pivotTable.ShowColumnGrandTotals = false;

                // Refresh and calculate the pivot table data using the correct API
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "HideColumnGrandTotalsDemo.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            HideColumnGrandTotalsDemo.Run();
        }
    }
}
