// Title: C# – Ascending numeric sort of a PivotTable row field using Aspose.Cells
// Description: Creates a workbook, adds Product and Sales data, builds a PivotTable, places Product as a row field and Sales as a data field, then enables auto‑sorting, sets ascending order, and uses the Sales field as the sort key before refreshing and saving the file.
// Keywords: Aspose.Cells PivotTable sort row field | C# numeric ascending sort PivotTable | auto sort pivot row by data field | Set AutoSortField Aspose.Cells | SortOrder Ascending PivotTable .NET | Excel pivot custom sort Aspose
// Common Searches: Aspose.Cells sort pivot row field by numeric value C# | how to enable auto sort for PivotTable row in Aspose.Cells | ascending sort of pivot row based on data column Aspose | C# code for custom numeric sorting in Excel pivot using Aspose | set AutoSortField index in Aspose.Cells PivotTable
// Developer Intent: The developer needs to order a PivotTable row field in ascending numeric order, using the values of a data field as the sort key, via Aspose.Cells for .NET.
// Use Cases: Generate sales reports where products are listed from lowest to highest revenue automatically. | Create dynamic Excel dashboards that re‑order rows whenever underlying numeric data changes. | Export Excel files with pre‑sorted PivotTables to avoid manual user sorting.
// AI Prompts: Show C# code to sort a PivotTable row field ascending by a numeric data field with Aspose.Cells. | Explain how to enable auto‑sorting and set the AutoSortField for a PivotTable row in Aspose.Cells. | Demonstrate changing the sort key to a different data column in an Aspose.Cells PivotTable.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotSortingDemo
{
    // Creates a workbook, adds Product and Sales data, builds a PivotTable, places Product as a row field and Sales as a data field, then enables auto‑sorting, sets ascending order, and uses the Sales field as the sort key before refreshing and saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            // Column A: Product, Column B: Sales (numeric)
            sheet.Cells["A1"].Value = "Product";
            sheet.Cells["B1"].Value = "Sales";
            sheet.Cells["A2"].Value = "B";
            sheet.Cells["A3"].Value = "A";
            sheet.Cells["A4"].Value = "C";
            sheet.Cells["B2"].Value = 200;   // numeric value
            sheet.Cells["B3"].Value = 300;   // numeric value
            sheet.Cells["B4"].Value = 100;   // numeric value

            // Add a pivot table covering the data range and place it at E3
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the product column as a row field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");

            // Add the sales column as a data field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Retrieve the row field we just added
            PivotField rowField = pivotTable.RowFields[0];

            // Enable automatic sorting and set it to ascend
            rowField.IsAutoSort = true;          // turn on auto‑sort
            rowField.IsAscendSort = true;        // sort ascending

            // Specify that sorting should be based on the first data field (Sales)
            // Index 0 refers to the first data field added to the pivot table
            rowField.AutoSortField = 0;

            // Alternatively, you can call SortBy directly:
            // rowField.SortBy(SortOrder.Ascending, 0);

            // Refresh the pivot table to apply the sorting
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_CustomNumericSort.xlsx");
        }
    }
}
