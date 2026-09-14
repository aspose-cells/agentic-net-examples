// Title: Apply a Top N filter to a row field of an Aspose.Cells pivot table using C#
// AI Prompts: Generate C# code that builds a pivot table with Aspose.Cells, adds a row field, and uses FilterTop10 to keep only the highest N items based on a sum data field. | Show how to limit a pivot table row field to the top 3 categories by total sales with Aspose.Cells for .NET. | Provide an example of programmatically applying a Top N filter to a pivot table row field and saving the workbook with Aspose.Cells.
// Common Searches: Aspose.Cells C# filter pivot table rows to top 5 values | How to use FilterTop10 on a row field in an Aspose.Cells pivot table | C# example applying a Top N filter to an Excel pivot table with Aspose.Cells | Show only highest categories in Aspose.Cells pivot table based on sum of sales | Apply Top 10 filter to pivot table row field programmatically in .NET
// Tags: Aspose.Cells pivot table top N filter | C# FilterTop10 row field | Aspose.Cells apply top values filter | pivot table row field sum filter | Excel workbook top categories Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample creates a workbook, fills it with category and sales data, defines a pivot table with Category as a row field and Sales as a data field, applies a Top N filter (configured for the top 3 items by sum of Sales) to the row field using FilterTop10, recalculates the pivot, and saves the file as PivotTop10Filter.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        // Header row
        sheet.Cells["A1"].Value = "Category";
        sheet.Cells["B1"].Value = "Sales";

        // Data rows
        string[] categories = { "Fruit", "Vegetable", "Fruit", "Dairy", "Vegetable", "Fruit", "Dairy", "Fruit" };
        int[] sales = { 120, 80, 150, 200, 90, 130, 110, 160 };

        for (int i = 0; i < categories.Length; i++)
        {
            sheet.Cells[i + 2, 0].Value = categories[i];
            sheet.Cells[i + 2, 1].Value = sales[i];
        }

        // Create a pivot table on the data range A1:B9, place it starting at D3
        int pivotIndex = sheet.PivotTables.Add("A1:B9", "D3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add "Category" as a row field (field index 0)
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);

        // Add "Sales" as a data field (field index 1)
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

        // Apply a Top 10 filter on the row field to show only the top 3 categories
        // Parameters: valueFieldIndex = 1 (Sales), type = Sum, isTop = true, itemCount = 3
        pivotTable.BaseFields[0].FilterTop10(1, PivotFilterType.Sum, true, 3);

        // Recalculate the pivot table to apply the filter
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("PivotTop10Filter.xlsx");
    }
}
