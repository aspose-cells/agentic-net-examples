// Title: How to enable drill‑down expand/collapse buttons on a PivotTable using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, adds sample data, builds a pivot table, and activates drill‑down with EnableDrilldown and ShowDrill properties in Aspose.Cells. | Show how to configure a PivotTable in Aspose.Cells to display the expand/collapse icons for aggregated cells. | Write a C# example that refreshes the pivot cache and calculates data after turning on drill‑down functionality.
// Common Searches: Aspose.Cells C# enable drilldown buttons on pivot table | Show expand collapse icons in Aspose.Cells pivot table example | C# Aspose.Cells pivot table EnableDrilldown property usage | How to display drill buttons for aggregated cells in Aspose.Cells .NET | Refresh pivot cache after setting ShowDrill in Aspose.Cells
// Tags: Aspose.Cells drilldown feature | Aspose.Cells expand collapse button setting | C# create pivot table Aspose.Cells | Aspose.Cells refresh pivot cache .NET | programmatic Excel pivot expand collapse buttons

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates creating a workbook, adding sample data, building a pivot table, enabling drill‑down and displaying expand/collapse buttons with EnableDrilldown and ShowDrill, refreshing and calculating the pivot, and saving the file as PivotTableShowDrillDemo.xlsx.
    public class PivotTableShowDrillDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and name it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Value";
            sheet.Cells["A2"].Value = "A";
            sheet.Cells["B2"].Value = 100;
            sheet.Cells["A3"].Value = "B";
            sheet.Cells["B3"].Value = 200;
            sheet.Cells["A4"].Value = "A";
            sheet.Cells["B4"].Value = 150;
            sheet.Cells["A5"].Value = "B";
            sheet.Cells["B5"].Value = 250;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

            // Enable drilldown functionality and show expand/collapse buttons
            pivotTable.EnableDrilldown = true; // Allows users to drill down into aggregated cells
            pivotTable.ShowDrill = true;       // Displays the expand/collapse (drill) buttons

            // Refresh the pivot cache data and calculate the results
            pivotTable.RefreshData();   // Correct method to refresh pivot cache
            pivotTable.CalculateData();

            // Save the workbook to a file
            string outputPath = "PivotTableShowDrillDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
