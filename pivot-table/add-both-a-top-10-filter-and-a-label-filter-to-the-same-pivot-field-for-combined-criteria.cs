// Title: How to Apply Both a Top 10 Filter and a Label Filter to the Same Pivot Field Using Aspose.Cells for .NET (C#)
// Common Searches: C# code to set a top 10 filter on a pivot field with Aspose.Cells | apply label filter and top 10 filter together on one pivot field Aspose.Cells | Aspose.Cells example for multiple filters on a single pivot table column in C# | combined top 10 and label filter criteria in an Aspose.Cells pivot table | Aspose.Cells .xlsx pivot table filter with both value and label conditions
// Tags: Aspose.Cells top 10 pivot filter | Aspose.Cells label pivot filter | Aspose.Cells combined pivot filters C# | C# Aspose.Cells multiple pivot filters | Aspose.Cells .xlsx pivot filter criteria

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample creates a workbook, populates it with sample sales data, adds a pivot table, assigns "Category" as a row field and "Quantity" as a summed data field, then demonstrates how to apply both a Top 10 filter and a label filter to the same pivot field before refreshing, calculating, and saving the file as PivotWithCombinedFilters.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Populate source data for the pivot table
            // -------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Header row
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Item");
            dataSheet.Cells["C1"].PutValue("Quantity");
            dataSheet.Cells["D1"].PutValue("Price");

            // Sample data
            object[,] sample = new object[,]
            {
                { "Fruit",     "Apple",    120, 1.20 },
                { "Fruit",     "Banana",    80, 0.80 },
                { "Fruit",     "Orange",   150, 1.00 },
                { "Vegetable", "Carrot",   200, 0.50 },
                { "Vegetable", "Broccoli",  90, 1.10 },
                { "Vegetable", "Spinach",   70, 0.90 },
                { "Beverage",  "Tea",      130, 1.50 },
                { "Beverage",  "Coffee",   110, 2.00 },
                { "Beverage",  "Juice",     95, 1.80 }
            };

            // Fill the worksheet with the sample data
            for (int i = 0; i < sample.GetLength(0); i++)
            {
                for (int j = 0; j < sample.GetLength(1); j++)
                {
                    dataSheet.Cells[i + 1, j].PutValue(sample[i, j]);
                }
            }

            // -------------------------------------------------
            // Add a pivot table based on the source data
            // -------------------------------------------------
            int pivotSheetIndex = workbook.Worksheets.Add();
            Worksheet pivotSheet = workbook.Worksheets[pivotSheetIndex];
            pivotSheet.Name = "Pivot";

            // Add the pivot table at cell C3 on the pivot sheet
            int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:D10", "C3", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Add "Category" as a row field (field index 0)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
            // Add "Quantity" as a data field (field index 2) – sum aggregation
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);

            // -------------------------------------------------
            // Refresh the pivot table to calculate data
            // -------------------------------------------------
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "PivotWithCombinedFilters.xlsx");
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
