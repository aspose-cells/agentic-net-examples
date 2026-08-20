// Title: Enable Drill‑Through (Drilldown) in Aspose.Cells PivotTable – C# Example
// Description: Creates a workbook, adds sample data, builds a PivotTable, places the Category field in rows and Amount in data, sets the EnableDrilldown property to true, refreshes the cache, calculates the pivot, and saves the file as EnableDrillthroughDemo.xlsx.
// Keywords: Aspose.Cells drilldown | EnableDrilldown property | C# pivot table drill‑through | Aspose.Cells RefreshData | Aspose.Cells CalculateData | interactive Excel PivotTable | Aspose.Cells .NET example
// Common Searches: Aspose.Cells enable drill‑through pivot table C# | How to set EnableDrilldown in Aspose.Cells | C# code to add drilldown to Aspose.Cells PivotTable | Refresh pivot cache after enabling drilldown Aspose.Cells | Calculate pivot data with EnableDrilldown Aspose.Cells
// Developer Intent: Add drill‑through capability to a PivotTable so users can double‑click a summary cell and view the underlying source rows.
// Use Cases: Interactive Excel reports where clicking a pivot cell reveals detailed records. | Workbooks that need to refresh their pivot cache while preserving drill‑through functionality. | Self‑service analytics dashboards built with Aspose.Cells that allow end‑users to explore raw data from aggregated views.
// AI Prompts: Generate C# code that creates a PivotTable with Aspose.Cells, enables drill‑through, refreshes the cache, calculates data, and saves the workbook. | Explain the steps required after setting EnableDrilldown to true in Aspose.Cells, including RefreshData and CalculateData. | Show an example of using EnableDrilldown together with RefreshData for an interactive pivot report in .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, builds a PivotTable, places the Category field in rows and Amount in data, sets the EnableDrilldown property to true, refreshes the cache, calculates the pivot, and saves the file as EnableDrillthroughDemo.xlsx.
    public class EnableDrillthroughDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].Value = "Category";
                sheet.Cells["B1"].Value = "Amount";
                sheet.Cells["A2"].Value = "Food";
                sheet.Cells["B2"].Value = 120;
                sheet.Cells["A3"].Value = "Drink";
                sheet.Cells["B3"].Value = 80;
                sheet.Cells["A4"].Value = "Food";
                sheet.Cells["B4"].Value = 150;

                // Add a pivot table covering the data range
                int ptIndex = sheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[ptIndex];

                // Configure pivot fields
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Enable drill‑through (drilldown) functionality
                pivot.EnableDrilldown = true;

                // Refresh the pivot cache to reflect the source data
                pivot.RefreshData();

                // Calculate the pivot table
                pivot.CalculateData();

                // Save the workbook
                workbook.Save("EnableDrillthroughDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            EnableDrillthroughDemo.Run();
        }
    }
}
