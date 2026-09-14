// Title: How to apply a CaptionEqual label filter to a row field in an Aspose.Cells pivot table using C#
// AI Prompts: Write C# code that creates a workbook, builds a pivot table from a range, and adds a label filter to show only rows where the Product field equals "Apple" using Aspose.Cells. | Show C# example that changes the pivot table label filter from an exact match to a contains condition for a row field in Aspose.Cells. | Provide C# sample that adds multiple label filters to a pivot table to include several product names (e.g., "Apple" and "Cherry") with Aspose.Cells.
// Common Searches: Aspose.Cells C# filter pivot table rows by specific value | How to apply a CaptionEqual filter to a pivot table using Aspose.Cells | Filter pivot table for multiple products in Aspose.Cells .NET | Refresh pivot table after setting label filter in Aspose.Cells C#
// Tags: Aspose.Cells pivot row filter | Exact text filter Aspose.Cells | C# pivot table row field filter | Update pivot cache Aspose.Cells | Aspose.Cells multiple row filters

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example creates a workbook, populates it with product and sales data, adds a pivot table, places the Product field in the row area, applies a CaptionEqual label filter so only rows with "Apple" appear, refreshes the pivot cache, calculates the data, and saves the workbook as PivotLabelFilterDemo.xlsx.
    public class PivotLabelFilterDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(80);
                sheet.Cells["A4"].PutValue("Apple");
                sheet.Cells["B4"].PutValue(150);
                sheet.Cells["A5"].PutValue("Cherry");
                sheet.Cells["B5"].PutValue(200);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Add the "Product" field to the row area and "Sales" to the data area
                pivot.AddFieldToArea(PivotFieldType.Row, "Product");
                pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Obtain the collection of pivot filters for the pivot table
                PivotFilterCollection filters = pivot.PivotFilters;

                // Add a label filter to show only rows where the product name equals "Apple"
                filters.AddLabelFilter(
                    baseFieldIndex: 0,
                    type: PivotFilterType.CaptionEqual,
                    label1: "Apple",
                    label2: null);

                // Refresh the pivot cache and calculate data
                pivot.RefreshData();      // correct API to refresh pivot cache
                pivot.CalculateData();

                // Save the workbook with the applied filter
                string outputPath = "PivotLabelFilterDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            PivotLabelFilterDemo.Run();
        }
    }
}
