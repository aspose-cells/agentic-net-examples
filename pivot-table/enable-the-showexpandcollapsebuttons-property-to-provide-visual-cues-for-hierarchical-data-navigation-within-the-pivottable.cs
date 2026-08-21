// Title: How to enable expand/collapse buttons for a PivotTable using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, adds sample data, builds a PivotTable, sets EnableDrilldown and ShowDrill to true, refreshes the cache, calculates the pivot data, and saves the file. | Update an existing Aspose.Cells PivotTable in C# to turn on drill‑down and display the hierarchy expand/collapse icons programmatically.
// Common Searches: Aspose.Cells C# show expand collapse icons in pivot table | Enable drilldown and ShowDrill property for PivotTable using Aspose.Cells .NET | Programmatically display hierarchy buttons in Aspose.Cells PivotTable | Refresh and calculate pivot table after enabling drilldown in C# Aspose.Cells | How to add expand/collapse buttons to a PivotTable with Aspose.Cells for .NET
// Tags: Aspose.Cells PivotTable enable drilldown | Aspose.Cells show expand collapse icons | C# Aspose.Cells refresh pivot cache | C# Aspose.Cells calculate pivot data | Aspose.Cells PivotTable hierarchical navigation

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates creating a workbook, populating data, adding a PivotTable, enabling drill‑down with visible expand/collapse buttons via the ShowDrill property, refreshing the cache, calculating the pivot data, and saving the workbook as an .xlsx file.
    public class PivotTableShowExpandCollapseDemo
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
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["A4"].PutValue("A");
            sheet.Cells["B4"].PutValue(150);
            sheet.Cells["A5"].PutValue("B");
            sheet.Cells["B5"].PutValue(250);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

            // Enable drilldown (allows expanding/collapsing)
            pivotTable.EnableDrilldown = true;

            // Show expand/collapse buttons in the pivot table
            pivotTable.ShowDrill = true;

            // Refresh and calculate the pivot table data using the correct API
            pivotTable.RefreshData();   // Refreshes the cache
            pivotTable.CalculateData(); // Calculates the pivot data

            // Save the workbook to a file
            try
            {
                workbook.Save("PivotTableShowExpandCollapseDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
