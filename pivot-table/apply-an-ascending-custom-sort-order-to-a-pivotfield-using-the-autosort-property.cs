// Title: Ascending Custom Sort for a PivotField using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample data, builds a pivot table, places the "Category" column as a row field, enables automatic sorting, sets the direction to ascending, and sorts by the field's own labels (AutoSortField = -1) before refreshing and saving the file.
// Keywords: Aspose.Cells | PivotField AutoSort | ascending sort C# | pivot table custom sort .NET | row field sorting | Excel export Aspose | IsAutoSort example | AutoSortField -1
// Common Searches: Aspose.Cells set ascending sort on PivotField | C# pivot table AutoSort example | How to enable AutoSort for row field in Aspose.Cells | Custom sort PivotField by label Aspose | Sort PivotTable rows alphabetically using Aspose.Cells
// Developer Intent: Enable a PivotField to sort its row labels automatically in ascending order by configuring IsAutoSort, IsAscendSort, and AutoSortField properties.
// Use Cases: Generate a sales dashboard where product categories appear alphabetically without manual re‑ordering. | Create a financial pivot that lists expense categories from A‑Z automatically after data refresh. | Export Excel reports that guarantee consistent row‑label ordering for downstream processing.
// AI Prompts: Write C# code with Aspose.Cells to apply a descending custom sort to a PivotField based on a numeric data field. | Show how to configure multiple row fields with independent AutoSort settings in a single pivot table using Aspose.Cells. | Explain how to turn off automatic sorting for a PivotField after it has been enabled in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotCustomSort
{
    // Creates a workbook, adds sample data, builds a pivot table, places the "Category" column as a row field, enables automatic sorting, sets the direction to ascending, and sorts by the field's own labels (AutoSortField = -1) before refreshing and saving the file.
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
            cells["B1"].Value = "Amount";
            cells["A2"].Value = "B";
            cells["A3"].Value = "A";
            cells["A4"].Value = "C";
            cells["B2"].Value = 200;
            cells["B3"].Value = 300;
            cells["B4"].Value = 100;

            // Add a pivot table based on the data range
            int ptIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[ptIndex];

            // Add the "Category" field as a row field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

            // Add the "Amount" field as a data field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Retrieve the row field we just added
            PivotField rowField = pivotTable.RowFields[0];

            // Enable automatic sorting for the field
            rowField.IsAutoSort = true;

            // Set the sort direction to ascending
            rowField.IsAscendSort = true;

            // Specify that the field should be sorted by its own labels (-1)
            rowField.AutoSortField = -1;

            // Refresh the pivot table data and calculate the results
            pivotTable.RefreshDataFlag = true;
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotFieldCustomAscendingSort.xlsx");
        }
    }
}
