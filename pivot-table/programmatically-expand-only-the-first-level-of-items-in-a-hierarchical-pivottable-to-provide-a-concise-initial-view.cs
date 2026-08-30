// Title: How to programmatically expand only the first‑level row items in an Aspose.Cells PivotTable using C#
// AI Prompts: Generate C# code with Aspose.Cells that creates a pivot table, collapses every row field, and then expands only the top‑level Category items. | Show how to use PivotField.HideDetail and PivotItem.IsDetailHidden properties in Aspose.Cells to control drill‑down visibility for hierarchical pivot tables. | Provide a complete example that saves the workbook after expanding the first row field while keeping deeper levels hidden.
// Common Searches: C# Aspose.Cells expand only top level rows in hierarchical pivot table | how to hide all pivot details and show only first row field Aspose.Cells | programmatically control drilldown visibility of pivot table rows using Aspose.Cells .NET | Aspose.Cells pivot table collapse all items then expand category level
// Tags: expand first level pivot rows Aspose.Cells C# | collapse all pivot items Aspose.Cells | pivotfield HideDetail Aspose.Cells | pivotitem IsDetailHidden Aspose.Cells | drilldown control hierarchical pivot Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExpandFirstLevel
{
    // The example creates a workbook with hierarchical data, builds a pivot table, collapses all row items, expands only the top‑level Category items using HideDetail and IsDetailHidden, and saves the result as PivotTableFirstLevelExpanded.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample hierarchical data
            // Category -> SubCategory -> Amount
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("SubCategory");
            sheet.Cells["C1"].PutValue("Amount");

            sheet.Cells["A2"].PutValue("Fruit");
            sheet.Cells["B2"].PutValue("Apple");
            sheet.Cells["C2"].PutValue(120);

            sheet.Cells["A3"].PutValue("Fruit");
            sheet.Cells["B3"].PutValue("Banana");
            sheet.Cells["C3"].PutValue(80);

            sheet.Cells["A4"].PutValue("Vegetable");
            sheet.Cells["B4"].PutValue("Carrot");
            sheet.Cells["C4"].PutValue(50);

            sheet.Cells["A5"].PutValue("Vegetable");
            sheet.Cells["B5"].PutValue("Potato");
            sheet.Cells["C5"].PutValue(70);

            // Add a new worksheet for the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

            // Create the pivot table
            int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:C5", "A3", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Add fields: first level (Category), second level (SubCategory), and data (Amount)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Row, "SubCategory");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Enable drilldown and show expand/collapse buttons
            pivotTable.EnableDrilldown = true;
            pivotTable.ShowDrill = true;

            // Refresh data so that pivot items are generated
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Collapse all levels by hiding details for each row field
            foreach (PivotField rowField in pivotTable.RowFields)
            {
                rowField.HideDetail(true);
            }

            // Expand only the first level (Category) items
            // The first row field (index 0) represents the top‑level items
            PivotField firstLevelField = pivotTable.RowFields[0];
            foreach (PivotItem item in firstLevelField.PivotItems)
            {
                // Ensure the detail of this top‑level item is visible
                item.IsDetailHidden = false;
            }

            // Save the workbook
            workbook.Save("PivotTableFirstLevelExpanded.xlsx");
        }
    }
}
