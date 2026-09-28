// Title: How to remove every filter from an Aspose.Cells pivot table in C# and refresh the workbook
// AI Prompts: Generate C# code that loops through each PivotTable field (rows, columns, pages) and invokes ClearFilter using Aspose.Cells. | Show how to empty the PivotFilters collection of a PivotTable and then recalculate the pivot data with Aspose.Cells. | Create a full example that removes every filter, refreshes the pivot, and saves the workbook to an .xlsx file using Aspose.Cells.
// Common Searches: asp.net core remove pivot table filters Aspose.Cells example | c# clear row field filter in Aspose.Cells pivot table | how to reset column filters in Aspose.Cells PivotTable | refresh pivot after clearing filters Aspose.Cells .NET | clear all pivot filters and recalculate data Aspose.Cells C#
// Tags: Aspose.Cells PivotTable ClearFilter C# | remove pivot filters Aspose.Cells .xlsx | refresh pivot data Aspose.Cells | iterate pivot fields Aspose.Cells | clear PivotFilters collection Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds sample data, builds a pivot table, applies label filters, then removes all filters by iterating over row, column, and page fields and clearing the PivotFilters collection. After clearing, it refreshes and recalculates the pivot table and saves the result as PivotTable_NoFilters.xlsx.
    class RemoveAllPivotFilters
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
                worksheet.Cells["C2"].PutValue(100);

                worksheet.Cells["A3"].PutValue("Fruit");
                worksheet.Cells["B3"].PutValue("Banana");
                worksheet.Cells["C3"].PutValue(150);

                worksheet.Cells["A4"].PutValue("Vegetable");
                worksheet.Cells["B4"].PutValue("Carrot");
                worksheet.Cells["C4"].PutValue(200);

                // Add a pivot table based on the data range
                int pivotIndex = worksheet.PivotTables.Add("A1:C4", "E1", "PivotTable1");
                PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

                // Configure the pivot table fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Apply some filters for demonstration purposes
                // Row field: show only "Fruit"
                pivotTable.RowFields[0].FilterByLabel(PivotFilterType.CaptionEqual, "Fruit", null);
                // Column field: show only "Apple"
                pivotTable.ColumnFields[0].FilterByLabel(PivotFilterType.CaptionEqual, "Apple", null);

                // Refresh the pivot table to apply the filters
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // -------------------------------------------------
                // Remove all filters from the pivot table
                // -------------------------------------------------

                // Clear filters on each row field
                foreach (PivotField field in pivotTable.RowFields)
                {
                    field.ClearFilter();
                }

                // Clear filters on each column field
                foreach (PivotField field in pivotTable.ColumnFields)
                {
                    field.ClearFilter();
                }

                // Clear filters on each page field (if any)
                foreach (PivotField field in pivotTable.PageFields)
                {
                    field.ClearFilter();
                }

                // Additionally clear any filters stored in the PivotFilters collection
                PivotFilterCollection pivotFilters = pivotTable.PivotFilters;
                int totalFieldCount = pivotTable.RowFields.Count + pivotTable.ColumnFields.Count + pivotTable.PageFields.Count;
                for (int i = 0; i < totalFieldCount; i++)
                {
                    pivotFilters.ClearFilter(i);
                }

                // Refresh the pivot table again to reflect the removal of filters
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook with the unfiltered pivot table
                workbook.Save("PivotTable_NoFilters.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            RemoveAllPivotFilters.Run();
        }
    }
}
