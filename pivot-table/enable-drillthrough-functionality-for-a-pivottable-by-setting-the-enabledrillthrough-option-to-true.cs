// Title: C# – Enable Drill‑Through (EnableDrilldown) for an Aspose.Cells PivotTable
// Description: Creates a workbook, adds sample data, inserts a PivotTable, assigns the Category field to rows and Amount to data, turns on the EnableDrilldown property, refreshes the pivot, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells PivotTable drill‑through C# | EnableDrilldown property .NET | Aspose.Cells pivot drilldown example | C# Excel pivot drill‑through | Aspose.Cells interactive reports
// Common Searches: how to enable drill‑through in Aspose.Cells PivotTable C# | Aspose.Cells EnableDrilldown usage | C# code for pivot table drill‑down with Aspose.Cells | Aspose.Cells pivot drill‑through sample | enable drill‑through for Excel pivot in .NET
// Developer Intent: Activate the drill‑through feature on a PivotTable so end‑users can double‑click a summary cell and see the underlying source rows in Excel.
// Use Cases: Interactive financial dashboards where users can explore transaction details from aggregated totals. | Sales reports that let analysts drill down from region‑level summaries to individual order rows. | Automated workbook generation with pivot tables that retain drill‑through capability when opened in Excel.
// AI Prompts: Generate C# code that builds a PivotTable with Aspose.Cells, enables drill‑through via EnableDrilldown, and saves the workbook. | Explain how to capture the rows returned by a drill‑through action from an Aspose.Cells PivotTable programmatically. | Show how to customize the target worksheet or formatting for drill‑through results after setting EnableDrilldown in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDrillthroughDemo
{
    // Creates a workbook, adds sample data, inserts a PivotTable, assigns the Category field to rows and Amount to data, turns on the EnableDrilldown property, refreshes the pivot, and saves the file as an Excel workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and add sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Amount";
            sheet.Cells["A2"].Value = "Food";
            sheet.Cells["B2"].Value = 120;
            sheet.Cells["A3"].Value = "Food";
            sheet.Cells["B3"].Value = 80;
            sheet.Cells["A4"].Value = "Drink";
            sheet.Cells["B4"].Value = 150;
            sheet.Cells["A5"].Value = "Drink";
            sheet.Cells["B5"].Value = 200;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Enable drill‑through (drilldown) functionality
            pivotTable.EnableDrilldown = true;

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("PivotTableDrillthroughDemo.xlsx");
        }
    }
}
