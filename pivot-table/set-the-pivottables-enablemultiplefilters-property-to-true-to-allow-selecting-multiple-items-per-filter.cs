// Title: Enable multi‑select filters on an Aspose.Cells PivotTable in C#
// AI Prompts: Generate C# code that creates a workbook, adds a pivot table from a data range, and activates multi‑select filtering by setting AllowMultipleFiltersPerField to true. | Show how to configure an Aspose.Cells PivotTable to permit users to select multiple items in a filter field within a .NET application.
// Common Searches: Aspose.Cells C# enable multi select filter on pivot table | AllowMultipleFiltersPerField property usage example Aspose.Cells | How to configure multiple filters per field in an Aspose.Cells pivot table | C# code sample for multi‑select pivot filters using Aspose.Cells | Aspose.Cells .NET pivot table multiple filter selections tutorial
// Tags: Aspose.Cells pivot table multiple filters | C# AllowMultipleFiltersPerField | Aspose.Cells enable multi‑select filter | pivot table filter configuration .NET | Aspose.Cells workbook pivot example

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotMultipleFilters
{
    // The example creates a new workbook, fills it with sample data, adds a pivot table on range A1:B5, assigns a row field and a data field, enables multi‑select filtering by setting AllowMultipleFiltersPerField to true, and saves the file as PivotTable_MultipleFilters.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Amount";
            sheet.Cells["A2"].Value = "Fruit";
            sheet.Cells["B2"].Value = 10;
            sheet.Cells["A3"].Value = "Vegetable";
            sheet.Cells["B3"].Value = 20;
            sheet.Cells["A4"].Value = "Fruit";
            sheet.Cells["B4"].Value = 15;
            sheet.Cells["A5"].Value = "Grain";
            sheet.Cells["B5"].Value = 5;

            // Add a pivot table that uses the data range A1:B5 and place it starting at D3
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table: Category as row field, Amount as data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Enable multiple filters per field so users can select multiple items in a filter
            pivotTable.AllowMultipleFiltersPerField = true;

            // Save the workbook to a file
            workbook.Save("PivotTable_MultipleFilters.xlsx");
        }
    }
}
