// Title: C# – Expand All Items in an Aspose.Cells PivotTable (Enable Drilldown & Refresh)
// Description: Demonstrates how to display every row item in an Aspose.Cells PivotTable by enabling drill‑down, showing drill buttons, and refreshing the pivot. Aspose.Cells has no direct ExpandAll method, so this configuration forces the pivot to render all items before saving the workbook.
// Keywords: Aspose.Cells expand all pivot items | C# PivotTable expand collapsed rows | enable drilldown Aspose.Cells | refresh pivot table show all items | Aspose.Cells PivotTable ExpandAll alternative | programmatic pivot expansion .NET | Aspose.Cells PivotTable drill buttons
// Common Searches: how to expand all items in Aspose.Cells PivotTable C# | Aspose.Cells ExpandAll method missing | show all rows in pivot table using Aspose.Cells | enable drilldown for full pivot visibility Aspose.Cells | refresh pivot table to display every category Aspose.Cells
// Developer Intent: Programmatically make every collapsed row in a PivotTable visible.
// Use Cases: Generate Excel reports where the pivot must list all categories without user interaction. | Create workbooks for downstream processing that require a fully expanded pivot view. | Automate workbook creation with refreshed pivot data and drill buttons pre‑enabled.
// AI Prompts: Write C# code with Aspose.Cells that expands all PivotTable items by enabling drilldown and refreshing the pivot. | Suggest an alternative to a non‑existent ExpandAll method for showing every pivot item in Aspose.Cells. | Explain how to verify that a saved workbook contains a fully expanded PivotTable using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to display every row item in an Aspose.Cells PivotTable by enabling drill‑down, showing drill buttons, and refreshing the pivot. Aspose.Cells has no direct ExpandAll method, so this configuration forces the pivot to render all items before saving the workbook.
    public class ExpandAllPivotItemsDemo
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

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["A4"].PutValue("A");
            sheet.Cells["B4"].PutValue(150);

            // Add a pivot table to the worksheet
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

            // Enable drilldown and show drill buttons (required for expand/collapse)
            pivotTable.EnableDrilldown = true;
            pivotTable.ShowDrill = true;

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Note: Aspose.Cells does not provide a direct ExpandAll method.
            // The pivot table will display all items when refreshed with drill buttons enabled.

            // Save the workbook with the expanded pivot table
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ExpandedPivotTable.xlsx");
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
    }
}
