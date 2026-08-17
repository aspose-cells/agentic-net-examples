// Title: Aspose.Cells for .NET: Enable Multiple Filters and Multi‑Select on a PivotTable (C# Example)
// Description: This C# sample creates a workbook, populates product‑sales data, adds a PivotTable, and demonstrates how to turn on AllowMultipleFiltersPerField and IsMultipleItemSelectionAllowed. It applies two value filters, refreshes the pivot, prints the configuration, and saves the file as PivotMultipleFiltersDemo.xlsx.
// Keywords: Aspose.Cells PivotTable multiple filters | AllowMultipleFiltersPerField C# | IsMultipleItemSelectionAllowed | Aspose.Cells value filter example | programmatic pivot filter testing | C# Aspose.Cells pivot demo | Excel PivotTable multi‑select API | unit test Aspose.Cells pivot
// Common Searches: how to enable multiple filters per field in Aspose.Cells | Aspose.Cells C# set page field multi‑item selection | add value greater than filter to Aspose.Cells pivot | verify pivot filter settings with Aspose.Cells | Aspose.Cells PivotTable example on GitHub
// Developer Intent: Turn on multi‑filter support and page‑field multi‑select in a PivotTable and confirm the settings programmatically.
// Use Cases: Build a sales dashboard where users can choose several categories at once. | Apply dynamic numeric thresholds to a pivot before exporting the report. | Automate validation of PivotTable filter configuration in CI pipelines.
// AI Prompts: Generate C# code that creates a PivotTable with AllowMultipleFiltersPerField and enables multi‑item selection on a page field using Aspose.Cells. | Explain how to programmatically verify that multiple value filters are applied correctly in an Aspose.Cells PivotTable. | Suggest ways to select multiple items in the page field via code and retrieve the filtered result set.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotMultipleFiltersDemo
{
    // This C# sample creates a workbook, populates product‑sales data, adds a PivotTable, and demonstrates how to turn on AllowMultipleFiltersPerField and IsMultipleItemSelectionAllowed. It applies two value filters, refreshes the pivot, prints the configuration, and saves the file as PivotMultipleFiltersDemo.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Category");
                sheet.Cells["C1"].PutValue("Sales");

                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue("Fruit");
                sheet.Cells["C2"].PutValue(1200);

                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue("Fruit");
                sheet.Cells["C3"].PutValue(800);

                sheet.Cells["A4"].PutValue("Carrot");
                sheet.Cells["B4"].PutValue("Vegetable");
                sheet.Cells["C4"].PutValue(1500);

                sheet.Cells["A5"].PutValue("Broccoli");
                sheet.Cells["B5"].PutValue("Vegetable");
                sheet.Cells["C5"].PutValue(700);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
                pivotTable.AddFieldToArea(PivotFieldType.Page, "Category");

                // Enable multiple filters per field and multiple item selection on the page field
                pivotTable.AllowMultipleFiltersPerField = true;
                PivotField pageField = pivotTable.PageFields[0];
                pageField.IsMultipleItemSelectionAllowed = true;

                // Add value filters (value2 is required by the API; set to 0 when not used)
                PivotFilter filter1 = pivotTable.PivotFilters.AddValueFilter(
                    pivotTable.DataFields[0].Position,
                    0,
                    PivotFilterType.ValueGreaterThan,
                    1000.0,
                    0.0);

                PivotFilter filter2 = pivotTable.PivotFilters.AddValueFilter(
                    pivotTable.DataFields[0].Position,
                    0,
                    PivotFilterType.ValueLessThan,
                    2000.0,
                    0.0);

                // Refresh pivot data and calculate results
                pivotTable.RefreshData(); // RefreshData is the correct method for PivotTable
                pivotTable.CalculateData();

                // Output verification information
                Console.WriteLine("AllowMultipleFiltersPerField: " + pivotTable.AllowMultipleFiltersPerField);
                Console.WriteLine("Page field IsMultipleItemSelectionAllowed: " + pageField.IsMultipleItemSelectionAllowed);
                Console.WriteLine("Number of PivotFilters applied: " + pivotTable.PivotFilters.Count);
                Console.WriteLine("Filter 1 - Type: " + filter1.FilterType + ", Value1: " + filter1.Value1);
                Console.WriteLine("Filter 2 - Type: " + filter2.FilterType + ", Value1: " + filter2.Value1);

                // Save the workbook
                workbook.Save("PivotMultipleFiltersDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
