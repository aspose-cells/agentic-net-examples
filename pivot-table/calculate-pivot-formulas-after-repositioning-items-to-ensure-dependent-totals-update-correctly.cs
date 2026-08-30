// Title: Reorder pivot row items and refresh calculated totals with Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to change the PositionInSameParentNode of specific pivot row items and then refresh the pivot table. | After reordering pivot items, call CalculateRange, RefreshData, and CalculateData to update calculated fields before saving the workbook.
// Common Searches: how to programmatically change the order of pivot row items using Aspose.Cells C# | refresh pivot table calculated fields after moving items in Aspose.Cells .NET | Aspose.Cells pivot table recalculate totals after repositioning items | set PositionInSameParentNode for pivot items in C# example | update all pivot tables in workbook after item reordering Aspose.Cells
// Tags: reorder pivot row items Aspose.Cells | update pivot totals after item move | modify pivot item position C# | calculate pivot data fields Aspose.Cells | recalculate all workbook pivots .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotUpdateDemo
{
    // The example creates a workbook, fills it with sample sales data, adds a pivot table with row, column, and data fields, defines a calculated field, calculates the pivot, reorders specific row items by setting PositionInSameParentNode, recalculates the pivot, refreshes all pivot tables in the workbook, and saves the result as PivotRepositionedUpdated.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];

                // Populate sample data
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("SubCategory");
                dataSheet.Cells["C1"].PutValue("Amount");

                string[] categories = { "Fruit", "Fruit", "Vegetable", "Vegetable", "Fruit", "Vegetable" };
                string[] subCategories = { "Apple", "Banana", "Carrot", "Broccoli", "Orange", "Spinach" };
                double[] amounts = { 120, 80, 150, 200, 90, 110 };

                for (int i = 0; i < categories.Length; i++)
                {
                    int row = i + 2;
                    dataSheet.Cells[$"A{row}"].PutValue(categories[i]);
                    dataSheet.Cells[$"B{row}"].PutValue(subCategories[i]);
                    dataSheet.Cells[$"C{row}"].PutValue(amounts[i]);
                }

                // Add a pivot table on a new sheet
                Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
                int pivotIndex = pivotSheet.PivotTables.Add("A1:C7", "E3", "SalesPivot");
                PivotTable pivot = pivotSheet.PivotTables[pivotIndex];

                // Configure pivot fields
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Column, "SubCategory");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Add a calculated field (simple reference to Amount; complex formulas like SUM() are not supported)
                pivot.AddCalculatedField("TotalAmount", "Amount", true);

                // Initial calculation
                pivot.CalculateRange();
                pivot.RefreshData();
                pivot.CalculateData();

                // Reposition some pivot items within the same parent (Category)
                // Move "Banana" to appear before "Apple" and "Orange" after "Apple"
                PivotItemCollection fruitItems = pivot.RowFields["Category"].PivotItems;
                if (fruitItems["Banana"] != null) fruitItems["Banana"].PositionInSameParentNode = 1;
                if (fruitItems["Apple"] != null) fruitItems["Apple"].PositionInSameParentNode = 2;
                if (fruitItems["Orange"] != null) fruitItems["Orange"].PositionInSameParentNode = 3;

                // Recalculate after repositioning
                pivot.CalculateRange();
                pivot.RefreshData();
                pivot.CalculateData();

                // Refresh all pivot tables in the workbook (useful if multiple pivots exist)
                workbook.Worksheets.RefreshPivotTables();

                // Save the workbook
                string outputPath = "PivotRepositionedUpdated.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
