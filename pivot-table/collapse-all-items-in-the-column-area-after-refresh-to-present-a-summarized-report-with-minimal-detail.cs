// Title: Collapse all column field items in an Aspose.Cells pivot table after refreshing data (C#)
// AI Prompts: Generate C# code that creates a pivot table with row and column fields, refreshes its data source, and collapses every column field item using Aspose.Cells. | Show how to use PivotField.PivotItems.HideAllDetail(true) to hide column details after calling PivotTable.RefreshData in a .NET workbook. | Provide a method that builds sample data, adds a pivot table, refreshes it, collapses the column area, and saves the workbook as an .xlsx file with Aspose.Cells.
// Common Searches: Aspose.Cells C# collapse column items after pivot refresh | How to hide column field details in a pivot table using Aspose.Cells for .NET | Programmatically collapse pivot table column area after RefreshData in C# | Aspose.Cells example to summarize pivot report by collapsing column fields | C# code to refresh pivot cache and collapse all column items with Aspose.Cells
// Tags: pivot column collapse using Aspose.Cells | refresh pivot cache programmatically C# | hide column field details Aspose.Cells | generate summarized pivot report .xlsx | Aspose.Cells column area collapse after refresh

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates creating a workbook, adding sample data, building a pivot table, refreshing its cache, collapsing all column field items, and saving the result as CollapsedColumnItems.xlsx using Aspose.Cells for .NET.
    public class CollapseColumnItemsAfterRefresh
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                // Columns: Category, SubCategory, Sales
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("SubCategory");
                sheet.Cells["C1"].PutValue("Sales");

                sheet.Cells["A2"].PutValue("Electronics");
                sheet.Cells["B2"].PutValue("Laptop");
                sheet.Cells["C2"].PutValue(1200);

                sheet.Cells["A3"].PutValue("Electronics");
                sheet.Cells["B3"].PutValue("Phone");
                sheet.Cells["C3"].PutValue(800);

                sheet.Cells["A4"].PutValue("Furniture");
                sheet.Cells["B4"].PutValue("Chair");
                sheet.Cells["C4"].PutValue(150);

                sheet.Cells["A5"].PutValue("Furniture");
                sheet.Cells["B5"].PutValue("Table");
                sheet.Cells["C5"].PutValue(300);

                // Add a pivot table to a new worksheet
                Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
                int pivotIndex = pivotSheet.PivotTables.Add("=Sheet1!A1:C5", "E3", "SalesPivot");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");          // Row field
                pivotTable.AddFieldToArea(PivotFieldType.Column, "SubCategory");   // Column field
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");           // Data field (sum)

                // Enable drill‑down buttons so collapse/expand UI is visible
                pivotTable.ShowDrill = true;

                // Refresh the pivot cache and recalculate the pivot table
                pivotTable.RefreshData();   // Correct API to refresh data source
                pivotTable.CalculateData();

                // Collapse all items in the column area
                foreach (PivotField columnField in pivotTable.ColumnFields)
                {
                    // Pass true to hide details (collapse)
                    columnField.PivotItems.HideAllDetail(true);
                }

                // Save the workbook
                workbook.Save("CollapsedColumnItems.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            CollapseColumnItemsAfterRefresh.Run();
        }
    }
}
