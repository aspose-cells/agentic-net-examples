// Title: How to apply an ascending custom sort to a PivotField in an Aspose.Cells PivotTable with C#
// AI Prompts: Generate C# code that builds a workbook, adds a pivot table, and configures a row field to sort its items in ascending order using the AutoSort properties of Aspose.Cells. | Demonstrate enabling CustomListSort on a PivotTable and setting IsAutoSort, IsAscendSort, and AutoSortField to achieve label‑based ascending sorting for a specific field in Aspose.Cells for .NET.
// Common Searches: asp.net aspose.cells set pivot field ascending sort using autosort | c# enable custom list sorting for pivot table aspose.cells | how to use IsAutoSort and IsAscendSort on PivotField in Aspose.Cells | ascending label sort for pivot row field with Aspose.Cells .NET | apply custom sort order to pivot table rows in Aspose.Cells C#
// Tags: Aspose.Cells pivot table AutoSort ascending | C# set IsAutoSort on PivotField | Enable CustomListSort in Aspose.Cells pivot | Sort pivot row field by labels .NET | Aspose.Cells generate pivot table with custom sort

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotCustomSort
{
    // Shows how to create a workbook, populate data, add a PivotTable, and apply an ascending custom sort to the Item row field by configuring IsAutoSort, IsAscendSort, and AutoSortField, while enabling CustomListSort before refreshing and saving the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Item";
            cells["C1"].Value = "Quantity";

            cells["A2"].Value = "Fruit";
            cells["B2"].Value = "Apple";
            cells["C2"].Value = 30;

            cells["A3"].Value = "Fruit";
            cells["B3"].Value = "Banana";
            cells["C3"].Value = 20;

            cells["A4"].Value = "Fruit";
            cells["B4"].Value = "Cherry";
            cells["C4"].Value = 15;

            cells["A5"].Value = "Vegetable";
            cells["B5"].Value = "Carrot";
            cells["C5"].Value = 25;

            cells["A6"].Value = "Vegetable";
            cells["B6"].Value = "Broccoli";
            cells["C6"].Value = 10;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C6", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            // Row field: Item (we will apply custom sort to this field)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Item");
            // Column field: Category
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Category");
            // Data field: Quantity (sum)
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

            // Get the row field that we want to sort
            PivotField itemField = pivotTable.RowFields[0];

            // Enable automatic sorting and set it to ascending order
            itemField.IsAutoSort = true;          // Turn on AutoSort
            itemField.IsAscendSort = true;        // Ascending order
            itemField.AutoSortField = -1;         // Sort by the field itself (labels)

            // Enable custom list sorting for the pivot table (allows custom order definitions)
            pivotTable.CustomListSort = true;

            // Refresh data and calculate the pivot table to apply sorting
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_CustomAscendingSort.xlsx");
        }
    }
}
