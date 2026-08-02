// Title: C# – Expand Only the First Level of a Hierarchical PivotTable with Aspose.Cells for .NET
// Description: This Aspose.Cells for .NET example creates a workbook with Category → SubCategory → Amount data, builds a hierarchical PivotTable, enables drill‑down, and uses the PivotField.HideDetail(true) method to collapse the SubCategory level so that only the top‑level Category rows are visible when the file is opened.
// Keywords: Aspose.Cells C# PivotTable first level | hide subcategory rows Aspose.Cells | expand first level hierarchical pivot | PivotField.HideDetail Aspose | drilldown collapse pivot .NET | Aspose.Cells sample code | Excel pivot hierarchy programmatic | C# Excel pivot collapse
// Common Searches: Aspose.Cells hide detail rows in pivot | C# expand only first level of pivot table | programmatically collapse subcategory in Aspose pivot | how to start a pivot table collapsed using Aspose.Cells | Aspose.Cells .NET pivot drilldown example
// Developer Intent: Programmatically collapse the SubCategory level of a hierarchical PivotTable so that only the top‑level Category rows are displayed initially.
// Use Cases: Generate a summary report that shows totals per Category without overwhelming users with SubCategory details. | Create an Excel dashboard where the pivot starts collapsed for a cleaner first view and lets users drill down on demand. | Export data to Excel from an application and deliver a workbook that opens with a concise, top‑level view.
// AI Prompts: Write C# code using Aspose.Cells to build a hierarchical PivotTable and hide the detail rows of the first row field. | Show how to set PivotField.HideDetail(true) after enabling drill‑down to collapse sub‑levels in a pivot table. | Provide a complete Aspose.Cells example that expands only the first level of a pivot hierarchy and saves the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExpandFirstLevel
{
    // This Aspose.Cells for .NET example creates a workbook with Category → SubCategory → Amount data, builds a hierarchical PivotTable, enables drill‑down, and uses the PivotField.HideDetail(true) method to collapse the SubCategory level so that only the top‑level Category rows are visible when the file is opened.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet (data source)
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate sample hierarchical data
            // Category -> SubCategory -> Amount
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("SubCategory");
            dataSheet.Cells["C1"].PutValue("Amount");

            dataSheet.Cells["A2"].PutValue("Fruit");
            dataSheet.Cells["B2"].PutValue("Apple");
            dataSheet.Cells["C2"].PutValue(120);

            dataSheet.Cells["A3"].PutValue("Fruit");
            dataSheet.Cells["B3"].PutValue("Banana");
            dataSheet.Cells["C3"].PutValue(80);

            dataSheet.Cells["A4"].PutValue("Vegetable");
            dataSheet.Cells["B4"].PutValue("Carrot");
            dataSheet.Cells["C4"].PutValue(50);

            dataSheet.Cells["A5"].PutValue("Vegetable");
            dataSheet.Cells["B5"].PutValue("Potato");
            dataSheet.Cells["C5"].PutValue(70);

            // Add a new worksheet to host the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

            // Add the pivot table (source range, destination cell, name)
            int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:C5", "A3", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Add fields: Category as row, SubCategory as row (to create hierarchy), Amount as data
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Row, "SubCategory");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Enable drilldown and show drill buttons (optional, improves UI)
            pivotTable.EnableDrilldown = true;
            pivotTable.ShowDrill = true;

            // Collapse details for the first row field (Category) so that only the first level is visible.
            // This hides the SubCategory items under each Category, achieving the desired concise view.
            PivotField categoryField = pivotTable.RowFields[0];
            categoryField.HideDetail(true);

            // Refresh data and calculate the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTableFirstLevelExpanded.xlsx");
        }
    }
}
