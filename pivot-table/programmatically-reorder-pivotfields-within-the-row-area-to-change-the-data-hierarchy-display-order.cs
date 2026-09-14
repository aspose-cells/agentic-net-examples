// Title: How to reorder row fields in an Aspose.Cells pivot table using C# to change hierarchy display order
// AI Prompts: Generate C# code that uses Aspose.Cells to move a pivot table row field from one position to another. | Show a snippet that changes the row field order in a pivot table from Category→SubCategory to SubCategory→Category with Aspose.Cells. | Provide an example of using the RowFields collection to reorder pivot row fields in a .NET workbook.
// Common Searches: Aspose.Cells C# reorder pivot table row fields programmatically | How to change the order of row fields in a pivot table using Aspose.Cells .NET | Move pivot table row field to a different position with Aspose.Cells API | Changing pivot hierarchy order in a C# Aspose.Cells workbook
// Tags: pivot row field reorder Aspose.Cells | pivot hierarchy change Aspose.Cells | reorder pivot fields programmatically | Aspose.Cells pivot table row positioning | C# pivot table field manipulation

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotFieldReorderDemo
{
    // The example creates a workbook, adds sample data, builds a pivot table with Category and SubCategory as row fields, then uses the RowFields collection's Move method to swap their positions so SubCategory appears before Category, refreshes and calculates the pivot, and saves the result as an Excel file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
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
            cells["C4"].Value = 150;

            cells["A5"].Value = "Vegetable";
            cells["B5"].Value = "Broccoli";
            cells["C5"].Value = 90;

            // Add a pivot table based on the data range
            PivotTableCollection pivotTables = sheet.PivotTables;
            int pivotIndex = pivotTables.Add("A1:C5", "E3", "PivotTable1");
            PivotTable pivotTable = pivotTables[pivotIndex];

            // Add two fields to the row area: Category (index 0) and SubCategory (index 1)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0); // Category
            pivotTable.AddFieldToArea(PivotFieldType.Row, 1); // SubCategory

            // Add the data field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2); // Amount

            // At this point the hierarchy is: Category -> SubCategory
            // To change the display order to SubCategory -> Category, move the fields
            // RowFields is a PivotFieldCollection; use its Move method (currPos, destPos)
            // Current positions are zero‑based: 0 = Category, 1 = SubCategory
            pivotTable.RowFields.Move(0, 1); // Move Category from position 0 to position 1

            // Refresh and calculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("ReorderedPivotTable.xlsx");
        }
    }
}
