// Title: Enable drill‑through (drilldown) for a PivotTable using Aspose.Cells in C#
// AI Prompts: Create a workbook, add sample data, insert a PivotTable, set EnableDrilldown to true, invoke RefreshData and CalculateData, then save the file with Aspose.Cells for .NET. | Update an existing PivotTable by turning on drill‑through, recompute its values, and export the workbook as an .xlsx file using C#. | Programmatically activate drilldown on a PivotTable, call its refresh and calculation methods, and write the result to disk with Aspose.Cells.
// Common Searches: aspnet example of enabling pivot table drillthrough with Aspose.Cells | c# code to turn on EnableDrilldown property for Aspose.Cells PivotTable | how to refresh and calculate pivot data after enabling drilldown in Aspose.Cells | saving a workbook with drill‑through enabled using Aspose.Cells for .NET | sample Aspose.Cells pivot table with drilldown functionality in C#
// Tags: Aspose.Cells pivot drilldown activation | C# configure pivot table drillthrough | invoke RefreshData and CalculateData on pivot | save workbook with drilldown enabled | populate source data for Aspose.Cells pivot

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDrillthroughDemo
{
    // Demonstrates creating a workbook, adding sample data, inserting a PivotTable, enabling drill‑through via the EnableDrilldown property, refreshing and calculating the pivot data, and saving the workbook as an .xlsx file using Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Get the first worksheet to hold source data
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate sample data for the pivot table
            dataSheet.Cells["A1"].Value = "Category";
            dataSheet.Cells["B1"].Value = "Amount";
            dataSheet.Cells["A2"].Value = "Food";
            dataSheet.Cells["B2"].Value = 120;
            dataSheet.Cells["A3"].Value = "Food";
            dataSheet.Cells["B3"].Value = 80;
            dataSheet.Cells["A4"].Value = "Beverage";
            dataSheet.Cells["B4"].Value = 150;
            dataSheet.Cells["A5"].Value = "Beverage";
            dataSheet.Cells["B5"].Value = 200;

            // Add a new worksheet that will contain the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

            // Add a pivot table based on the source range A1:B5
            // Destination top‑left cell is C3
            int pivotIndex = pivotSheet.PivotTables.Add("Data!A1:B5", "C3", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Enable drill‑through (drilldown) functionality
            pivotTable.EnableDrilldown = true;

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook (lifecycle save)
            workbook.Save("PivotTableDrillthroughDemo.xlsx");

            // Optional: output confirmation
            Console.WriteLine("Workbook saved with drill‑through enabled.");
        }
    }
}
