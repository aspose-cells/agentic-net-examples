// Title: Programmatically select multiple items in an Aspose.Cells slicer, refresh it, and save the workbook as XLSX (C#)
// AI Prompts: Create a pivot table, add a slicer for the 'Category' field, set the slicer cache items 'Fruit' and 'Vegetable' to selected, refresh the slicer, and save the workbook to 'SlicerMultipleSelection.xlsx' using Aspose.Cells for .NET. | Retrieve an existing slicer from a worksheet, programmatically mark specific SlicerCacheItem values as selected, call Refresh on the slicer, and export the updated workbook to XLSX with the Aspose.Cells C# API.
// Common Searches: Aspose.Cells C# select multiple slicer items programmatically | Refresh slicer after changing selections with Aspose.Cells .NET | Save workbook after updating slicer cache items in C# | Retrieve slicer object from worksheet using Aspose.Cells API | How to set SlicerCacheItem.Selected property in Aspose.Cells
// Tags: aspnet slicer cache item selection | aspnet pivot table slicer refresh | aspnet save workbook with updated slicer | aspnet programmatic slicer manipulation | aspnet multiple slicer items selection

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// The example creates a workbook, adds sample data, builds a pivot table, inserts a slicer linked to the 'Category' field, programmatically selects the 'Fruit' and 'Vegetable' items by setting SlicerCacheItem.Selected, refreshes the slicer, and saves the file as SlicerMultipleSelection.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Amount");
        sheet.Cells["A2"].PutValue("Fruit");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["A3"].PutValue("Vegetable");
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["A4"].PutValue("Fruit");
        sheet.Cells["B4"].PutValue(200);

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a slicer linked to the pivot table for the "Category" field
        int slicerIndex = sheet.Slicers.Add(pivot, "F1", "Category");
        Slicer slicer = sheet.Slicers[slicerIndex];

        // Programmatically select multiple items (e.g., "Fruit" and "Vegetable")
        foreach (SlicerCacheItem item in slicer.SlicerCache.SlicerCacheItems)
        {
            // Select the desired items, deselect the rest
            if (item.Value == "Fruit" || item.Value == "Vegetable")
                item.Selected = true;
            else
                item.Selected = false;
        }

        // Refresh the slicer to apply the selection changes
        slicer.Refresh();

        // Save the workbook with the updated slicer selections
        workbook.Save("SlicerMultipleSelection.xlsx");
    }
}
