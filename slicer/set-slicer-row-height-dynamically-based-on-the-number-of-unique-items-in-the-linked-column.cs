// Title: Set Aspose.Cells slicer row height dynamically in C# using the count of unique items in the linked column
// AI Prompts: Create a workbook, add a pivot table, insert a slicer linked to a column, compute the number of distinct values in that column, and set slicer.RowHeight based on a base height plus an increment per extra item. | Resize the slicer's Shape.HeightPt to equal the calculated row height multiplied by the total distinct items so the slicer fully displays all rows. | Change the base row height and per‑item increment values to customize dynamic slicer sizing for different data sets in an Aspose.Cells .NET workbook.
// Common Searches: Aspose.Cells C# set slicer row height based on distinct column values | How to count unique items in a worksheet column for slicer sizing with Aspose.Cells | Programmatically adjust slicer shape height to fit all rows in a .NET Excel file | Dynamic slicer height example using Aspose.Cells pivot table | Adjust slicer RowHeight property after counting distinct categories in C#
// Tags: aspocells slicer rowheight dynamic | distinct column value count aspocells | pivot table linked slicer sizing | slicer shape height adjustment aspocells | c# dynamic slicer sizing based on unique items

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// Demonstrates creating a workbook, adding a pivot table, inserting a slicer linked to the Category field, counting distinct Category values, and dynamically setting both the slicer's RowHeight and overall shape height before saving the file.
class SlicerDynamicRowHeight
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data (Category column will be used for the slicer)
        cells["A1"].Value = "Category";
        cells["B1"].Value = "Value";
        cells["A2"].Value = "Fruit";
        cells["B2"].Value = 10;
        cells["A3"].Value = "Vegetable";
        cells["B3"].Value = 20;
        cells["A4"].Value = "Fruit";
        cells["B4"].Value = 15;
        cells["A5"].Value = "Grain";
        cells["B5"].Value = 5;
        cells["A6"].Value = "Vegetable";
        cells["B6"].Value = 25;

        // Add a pivot table based on the data range
        int pivotIdx = sheet.PivotTables.Add("A1:B6", "D3", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Data, "Value");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a slicer linked to the "Category" field of the pivot table
        int slicerIdx = sheet.Slicers.Add(pivot, "F3", "Category");
        Slicer slicer = sheet.Slicers[slicerIdx];

        // Determine the number of unique items in the slicer source column (Category)
        int firstDataRow = 1; // zero‑based index (row 2 in Excel)
        int lastDataRow = sheet.Cells.MaxDataRow; // last row with data
        HashSet<string> uniqueItems = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (int r = firstDataRow; r <= lastDataRow; r++)
        {
            string val = sheet.Cells[r, 0].StringValue; // column A (index 0)
            if (!string.IsNullOrEmpty(val))
                uniqueItems.Add(val);
        }
        int uniqueCount = uniqueItems.Count;

        // Set the row height dynamically: larger height for more items
        // Example rule: base height 18 pt, add 2 pt for each additional unique item beyond 3
        double baseHeight = 18.0;
        double extraPerItem = 2.0;
        double rowHeight = uniqueCount > 3 ? baseHeight + (uniqueCount - 3) * extraPerItem : baseHeight;
        slicer.RowHeight = rowHeight;

        // Optionally adjust the overall slicer height to fit all rows
        slicer.Shape.HeightPt = rowHeight * uniqueCount;

        // Save the workbook
        workbook.Save("SlicerDynamicRowHeight.xlsx");
    }
}
