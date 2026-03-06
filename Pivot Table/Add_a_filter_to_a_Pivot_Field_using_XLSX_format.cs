using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotFilterDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Amount";

            cells["A2"].Value = "Apple";
            cells["B2"].Value = 120;

            cells["A3"].Value = "Banana";
            cells["B3"].Value = 80;

            cells["A4"].Value = "Avocado";
            cells["B4"].Value = 150;

            cells["A5"].Value = "Cherry";
            cells["B5"].Value = 60;

            // Add a pivot table (create rule)
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");   // Row field (base field index 0)
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");    // Data field (value field index 1)

            // Add a label filter to the row field: show only categories that begin with "A"
            // Use the AddLabelFilter method from PivotFilterCollection (feature rule)
            PivotFilterCollection filters = pivot.PivotFilters;
            PivotFilter labelFilter = filters.AddLabelFilter(
                baseFieldIndex: 0,                     // Index of the "Category" field
                type: PivotFilterType.CaptionBeginsWith,
                label1: "A",                           // Filter condition
                label2: null);                         // Not used for this filter type

            // Optional: set a friendly name for the filter
            labelFilter.Name = "CategoriesStartingWithA";

            // Refresh and calculate the pivot table to apply the filter
            pivot.RefreshData();
            pivot.CalculateData();

            // Save the workbook in XLSX format (save rule)
            workbook.Save("PivotFieldFilterDemo.xlsx");
        }
    }
}