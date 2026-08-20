// Title: Clear All Filters from an Aspose.Cells PivotTable in C# (.NET)
// Description: Demonstrates how to create a workbook, add a PivotTable, apply label filters, then remove every filter (row, column, page, and PivotFilters collection) using ClearFilter, refresh the data, and save the file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells clear pivot filters | C# remove pivot table filters | Aspose.Cells PivotField ClearFilter | reset Aspose pivot filters programmatically | Aspose.Cells PivotFilters collection clear | unfilter Aspose.Cells PivotTable | Aspose.Cells .NET pivot table example
// Common Searches: how to clear all filters Aspose.Cells pivot table C# | remove row and column filters from Aspose pivot table | reset pivot table filters Aspose.Cells .NET | Aspose.Cells ClearFilter example | unfilter PivotTable using Aspose.Cells
// Developer Intent: Programmatically remove every filter applied to a PivotTable so the full dataset is displayed.
// Use Cases: Return a PivotTable to its original state after temporary analysis filters. | Ensure exported reports contain unfiltered data for all recipients. | Automate workbook preparation where no residual filters should remain.
// AI Prompts: Generate C# code with Aspose.Cells that clears all filters from a PivotTable, including row, column, page, and PivotFilters collections, then refreshes and saves the workbook. | Show an example of adding a PivotTable, applying label filters, removing them using ClearFilter, and explaining the required RefreshData and CalculateData calls. | Explain the difference between PivotField.ClearFilter and PivotFiltersCollection.ClearFilter in Aspose.Cells and when each should be used.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a PivotTable, apply label filters, then remove every filter (row, column, page, and PivotFilters collection) using ClearFilter, refresh the data, and save the file with Aspose.Cells for .NET.
    public class RemoveAllPivotFiltersDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Product");
                worksheet.Cells["C1"].PutValue("Sales");

                worksheet.Cells["A2"].PutValue("Fruit");
                worksheet.Cells["B2"].PutValue("Apple");
                worksheet.Cells["C2"].PutValue(120);

                worksheet.Cells["A3"].PutValue("Fruit");
                worksheet.Cells["B3"].PutValue("Banana");
                worksheet.Cells["C3"].PutValue(80);

                worksheet.Cells["A4"].PutValue("Vegetable");
                worksheet.Cells["B4"].PutValue("Carrot");
                worksheet.Cells["C4"].PutValue(150);

                worksheet.Cells["A5"].PutValue("Vegetable");
                worksheet.Cells["B5"].PutValue("Broccoli");
                worksheet.Cells["C5"].PutValue(130);

                // Add a pivot table based on the data range
                int pivotIndex = worksheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
                PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

                // Configure the pivot table fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Apply some filters to demonstrate clearing them later
                // Filter Row field "Category" to show only "Fruit"
                PivotField rowField = pivotTable.RowFields["Category"];
                rowField.FilterByLabel(PivotFilterType.CaptionEqual, "Fruit", null);

                // Filter Column field "Product" to show only "Apple"
                PivotField columnField = pivotTable.ColumnFields["Product"];
                columnField.FilterByLabel(PivotFilterType.CaptionEqual, "Apple", null);

                // Refresh the pivot table to reflect the applied filters
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // ------------------- Remove all filters -------------------
                // Clear filters on Row fields
                foreach (PivotField pf in pivotTable.RowFields)
                {
                    pf.ClearFilter();
                }

                // Clear filters on Column fields
                foreach (PivotField pf in pivotTable.ColumnFields)
                {
                    pf.ClearFilter();
                }

                // Clear filters on Page fields (if any)
                foreach (PivotField pf in pivotTable.PageFields)
                {
                    pf.ClearFilter();
                }

                // Additionally, clear any filters stored in the PivotFilters collection
                PivotFilterCollection filters = pivotTable.PivotFilters;
                for (int i = 0; i < filters.Count; i++)
                {
                    // The field index corresponds to the base field index of the filter
                    filters.ClearFilter(i);
                }

                // Refresh the pivot table again to show unfiltered data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "RemoveAllPivotFiltersDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            RemoveAllPivotFiltersDemo.Run();
        }
    }
}
