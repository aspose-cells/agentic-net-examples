// Title: C# Example: Expand First Level of a Hierarchical PivotTable with Aspose.Cells for .NET
// Description: This Aspose.Cells for .NET sample builds a workbook containing Category and SubCategory data, creates a PivotTable on a separate sheet, enables drill‑down buttons, applies outline layout, and then programmatically collapses all sub‑level rows so that only the top‑level Category items remain expanded. The pivot is refreshed, calculated, and saved as PivotTableFirstLevelExpanded.xlsx.
// Keywords: Aspose.Cells | C# | .NET | PivotTable | hierarchical pivot | first level expand | collapse detail rows | HideDetail | ShowDrill | outline form | row field API | Excel automation | sample code | GitHub example | Excel report generation
// Common Searches: Aspose.Cells expand first level of hierarchical pivot table C# | How to hide subcategory rows in Aspose.Cells PivotTable | Show drill‑down buttons but collapse detail rows in .NET PivotTable | Programmatically collapse PivotTable row field using Aspose.Cells | C# code to display only top‑level rows in an Excel pivot
// Developer Intent: Display only the top‑level row items of a hierarchical PivotTable while keeping drill‑down controls active.
// Use Cases: Create a summary workbook that initially shows only categories, letting users expand subcategories on demand. | Build a fast‑loading dashboard where detailed rows are hidden until the viewer chooses to explore them. | Generate a printable pivot view that lists primary groups without inner items for a clean layout.
// AI Prompts: Provide C# code that expands only the first level of a hierarchical PivotTable using Aspose.Cells and keeps drill‑down buttons visible. | Show how to programmatically collapse all detail rows of a specific row field in an Aspose.Cells PivotTable. | Explain how to toggle sub‑level visibility in an Aspose.Cells PivotTable after refreshing the data.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExpandFirstLevel
{
    // This Aspose.Cells for .NET sample builds a workbook containing Category and SubCategory data, creates a PivotTable on a separate sheet, enables drill‑down buttons, applies outline layout, and then programmatically collapses all sub‑level rows so that only the top‑level Category items remain expanded. The pivot is refreshed, calculated, and saved as PivotTableFirstLevelExpanded.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample hierarchical data
            // Row fields: Category (Level 1) and SubCategory (Level 2)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("SubCategory");
            sheet.Cells["C1"].PutValue("Amount");

            sheet.Cells["A2"].PutValue("Food");
            sheet.Cells["B2"].PutValue("Fruit");
            sheet.Cells["C2"].PutValue(120);

            sheet.Cells["A3"].PutValue("Food");
            sheet.Cells["B3"].PutValue("Vegetable");
            sheet.Cells["C3"].PutValue(80);

            sheet.Cells["A4"].PutValue("Beverage");
            sheet.Cells["B4"].PutValue("Tea");
            sheet.Cells["C4"].PutValue(50);

            sheet.Cells["A5"].PutValue("Beverage");
            sheet.Cells["B5"].PutValue("Coffee");
            sheet.Cells["C5"].PutValue(70);

            // Add a new worksheet for the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

            // Create the pivot table using the data range
            int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:C5", "A3", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Add hierarchical row fields (Category -> SubCategory)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Row, "SubCategory");

            // Add a data field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Show expand/collapse buttons
            pivotTable.ShowDrill = true;
            // Layout the pivot in outline form (makes hierarchy visible)
            pivotTable.ShowInOutlineForm();

            // Collapse all detail under the first row field (Category)
            // This leaves only the first level (Category) expanded.
            PivotField firstRowField = pivotTable.RowFields[0];
            firstRowField.HideDetail(true);

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTableFirstLevelExpanded.xlsx");
        }
    }
}
