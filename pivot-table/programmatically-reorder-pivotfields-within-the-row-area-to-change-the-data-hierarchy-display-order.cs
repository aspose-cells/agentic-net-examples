// Title: Reorder Pivot Row Fields Programmatically with Aspose.Cells for .NET
// Description: Shows how to modify the row‑field hierarchy in an Aspose.Cells pivot table by using the RowFields.Move method. The sample creates a workbook, builds a pivot table with Category and SubCategory, swaps their positions, refreshes the data, and saves the workbook.
// Keywords: Aspose.Cells pivot row field order | RowFields.Move C# | change pivot hierarchy .NET | programmatic pivot field reorder | Aspose.Cells PivotTable example | C# reorder pivot rows | Excel pivot table field position | Aspose.Cells API RowFields | move pivot field position | pivot table hierarchy Aspose
// Common Searches: Aspose.Cells move pivot row field | C# reorder pivot table fields | change row field order in Aspose.Cells | how to swap pivot fields programmatically | RowFields.Move usage example | Aspose.Cells pivot hierarchy change | set pivot table field sequence .NET
// Developer Intent: Change the order of row fields in a pivot table to modify the displayed hierarchy.
// Use Cases: Generate reports where SubCategory appears before Category. | Standardize pivot layouts across multiple workbooks before distribution. | Allow users to reorder fields dynamically in a web application via code. | Automate data‑analysis pipelines that require a specific pivot field sequence. | Create Excel templates that enforce a predefined row‑field order.
// AI Prompts: Write C# code that moves a specified pivot row field to a target index using Aspose.Cells. | Explain the steps to reorder pivot row fields and why RefreshData and CalculateData are required. | Generate a function that accepts a list of field names and reorders the pivot table rows accordingly. | Show how to verify the new field order after calling RowFields.Move. | Provide a PowerShell script that uses Aspose.Cells to reorder pivot fields in existing Excel files.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotFieldReorderDemo
{
    // Shows how to modify the row‑field hierarchy in an Aspose.Cells pivot table by using the RowFields.Move method. The sample creates a workbook, builds a pivot table with Category and SubCategory, swaps their positions, refreshes the data, and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            Cells cells = sheet.Cells;
            cells["A1"].Value = "Category";
            cells["B1"].Value = "SubCategory";
            cells["C1"].Value = "Amount";

            cells["A2"].Value = "Fruit";
            cells["B2"].Value = "Apple";
            cells["C2"].Value = 120;

            cells["A3"].Value = "Fruit";
            cells["B3"].Value = "Banana";
            cells["C3"].Value = 80;

            cells["A4"].Value = "Vegetable";
            cells["B4"].Value = "Carrot";
            cells["C4"].Value = 50;

            cells["A5"].Value = "Vegetable";
            cells["B5"].Value = "Broccoli";
            cells["C5"].Value = 70;

            // Add a pivot table based on the data range
            PivotTableCollection pivotTables = sheet.PivotTables;
            int pivotIndex = pivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = pivotTables[pivotIndex];

            // Add two fields to the row area: first Category, then SubCategory
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");      // position 0
            pivotTable.AddFieldToArea(PivotFieldType.Row, "SubCategory"); // position 1

            // Add the Amount field to the data area
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // At this point the hierarchy is Category > SubCategory.
            // To change the display order to SubCategory > Category, move the field.
            // Current positions: Category (0), SubCategory (1)
            // Move SubCategory (currPos = 1) to position 0.
            pivotTable.RowFields.Move(1, 0);

            // Refresh and calculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotFieldReorderResult.xlsx");
        }
    }
}
