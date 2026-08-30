// Title: Enable drill‑down and expand/collapse buttons on an Aspose.Cells pivot table in C#
// AI Prompts: Write C# code that builds a pivot table from a worksheet range and turns on the EnableDrilldown and ShowDrill properties using Aspose.Cells. | Adapt the example to load data from an existing Excel file, then configure the pivot table to allow drill‑down while keeping the same row and column fields. | Create a reusable method that receives a worksheet and a data range, adds a pivot table, and sets up drill‑down and expand/collapse UI elements for Aspose.Cells .NET.
// Common Searches: Aspose.Cells C# how to turn on drilldown for pivot table cells | show expand collapse buttons in Aspose.Cells pivot table example | double‑click a pivot cell to view source rows using Aspose.Cells .NET | enable drill‑down property on Aspose.Cells pivot table programmatically | C# code sample for creating pivot table with drill‑down in Aspose.Cells
// Tags: Aspose.Cells pivot drilldown configuration | C# Aspose.Cells pivot expand collapse UI | Aspose.Cells create pivot from cell range | Aspose.Cells refresh and calculate pivot data | Aspose.Cells enable drilldown property example

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsDrilldownDemo
{
    // Demonstrates creating a workbook, populating sample data, adding a pivot table, enabling drill‑down and expand/collapse buttons, refreshing the pivot cache, calculating results, and saving the file as an .xlsx document using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            // Columns: Category, SubCategory, Sales
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "SubCategory";
            sheet.Cells["C1"].Value = "Sales";

            sheet.Cells["A2"].Value = "Fruit";
            sheet.Cells["B2"].Value = "Apple";
            sheet.Cells["C2"].Value = 120;

            sheet.Cells["A3"].Value = "Fruit";
            sheet.Cells["B3"].Value = "Orange";
            sheet.Cells["C3"].Value = 150;

            sheet.Cells["A4"].Value = "Vegetable";
            sheet.Cells["B4"].Value = "Carrot";
            sheet.Cells["C4"].Value = 80;

            sheet.Cells["A5"].Value = "Vegetable";
            sheet.Cells["B5"].Value = "Broccoli";
            sheet.Cells["C5"].Value = 95;

            // Add a pivot table based on the data range
            // Destination top‑left cell is D3
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "D3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the pivot table:
            // - Row field: Category
            // - Column field: SubCategory
            // - Data field: Sales (sum)
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Column, "SubCategory");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Enable drill‑down functionality so users can double‑click a cell
            // and see the underlying records.
            pivot.EnableDrilldown = true;

            // Show the expand/collapse (drill) buttons in the UI.
            pivot.ShowDrill = true;

            // Refresh the pivot cache and calculate the results.
            pivot.RefreshData();
            pivot.CalculateData();

            // Save the workbook to a file.
            workbook.Save("PivotTableDrilldownDemo.xlsx");
        }
    }
}
