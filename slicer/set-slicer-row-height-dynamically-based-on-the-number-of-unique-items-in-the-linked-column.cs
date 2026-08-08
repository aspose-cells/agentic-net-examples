// Title: Aspose.Cells C# – Dynamically Set Slicer Row Height from Unique Column Values
// Description: This example creates a workbook, fills a Category column, builds a pivot table, adds a slicer linked to the Category field, counts distinct categories with a HashSet, and programmatically sets the slicer’s RowHeight based on that count before saving the file.
// Keywords: Aspose.Cells slicer height | C# dynamic slicer size | adjust slicer row height programmatically | unique values count Aspose.Cells | pivot table slicer .NET
// Common Searches: change slicer height based on distinct values Aspose.Cells | set slicer row height dynamically C# | calculate slicer size from unique items | Aspose.Cells adjust slicer dimensions automatically
// Developer Intent: Automatically resize a slicer’s row height to match the number of distinct entries in its source column.
// Use Cases: Generate reports where the slicer expands to display all filter options without scrolling. | Create templates that adapt slicer dimensions when data sets vary in size. | Build dashboards that maintain a consistent layout regardless of the number of categories.
// AI Prompts: Show C# code using Aspose.Cells to compute the count of unique values in a worksheet column and set slicer.RowHeight accordingly. | Provide a step‑by‑step example that extracts distinct items, multiplies by a base height, and applies the result to a slicer. | Explain how to link a slicer to a pivot field and programmatically adjust its height based on data diversity in Aspose.Cells for .NET.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// This example creates a workbook, fills a Category column, builds a pivot table, adds a slicer linked to the Category field, counts distinct categories with a HashSet, and programmatically sets the slicer’s RowHeight based on that count before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data (Category column will be used for the slicer)
        cells["A1"].Value = "Category";
        cells["B1"].Value = "Amount";

        string[] categories = { "Fruit", "Fruit", "Vegetable", "Fruit", "Grain", "Vegetable" };
        int[] amounts = { 10, 20, 15, 30, 5, 12 };

        for (int i = 0; i < categories.Length; i++)
        {
            cells[i + 2, 0].Value = categories[i];   // Column A
            cells[i + 2, 1].Value = amounts[i];     // Column B
        }

        // Create a pivot table based on the data range and place it at D3
        int pivotIdx = sheet.PivotTables.Add("A1:B7", "D3", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a slicer linked to the "Category" field, positioned at F3
        int slicerIdx = sheet.Slicers.Add(pivot, "F3", "Category");
        Slicer slicer = sheet.Slicers[slicerIdx];

        // Calculate the number of unique items in the linked column (Category)
        HashSet<string> uniqueItems = new HashSet<string>();
        for (int row = 1; row <= cells.MaxDataRow; row++) // start from row 2 (index 1)
        {
            object val = cells[row, 0].Value; // Column A
            if (val != null)
                uniqueItems.Add(val.ToString());
        }
        int uniqueCount = uniqueItems.Count;

        // Dynamically set the slicer row height (example: 15 points per unique item)
        double baseHeight = 15.0;
        slicer.RowHeight = baseHeight * uniqueCount;

        // Save the workbook
        workbook.Save("SlicerDynamicRowHeight.xlsx");
    }
}
