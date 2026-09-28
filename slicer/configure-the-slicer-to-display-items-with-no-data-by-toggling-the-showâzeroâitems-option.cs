// Title: How to configure an Aspose.Cells slicer to show zero‑value items in a pivot table with C#
// AI Prompts: Write C# code that creates a workbook, adds a pivot table, inserts a slicer for the "Category" field, and sets ShowAllItems = true and ShowTypeOfItemsWithNoData = Natural so that categories with zero values appear in the slicer. | Modify an existing Aspose.Cells slicer in a .NET project to enable the show‑zero‑items option and refresh it, ensuring items without data are selectable. | Generate an Excel file using Aspose.Cells where the slicer linked to a pivot table lists a category with a zero value as a selectable item, demonstrating ItemsWithNoDataShowMode usage.
// Common Searches: Aspose.Cells C# slicer show items with no data in pivot table | how to enable show zero items option for slicer using Aspose.Cells | C# example displaying categories with zero values in Excel slicer | Aspose.Cells set ShowAllItems and ItemsWithNoDataShowMode for slicer | display empty categories in pivot slicer using Aspose.Cells .NET
// Tags: Aspose.Cells slicer ShowAllItems property | Aspose.Cells ItemsWithNoDataShowMode usage | C# pivot table slicer zero-value items | Excel slicer display no-data categories Aspose.Cells | configure slicer show zero items .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// The example creates a workbook, populates sample data, builds a pivot table, adds a slicer linked to the "Category" field, enables ShowAllItems and sets ShowTypeOfItemsWithNoData to Natural so that categories with zero values (e.g., "D") are shown in the slicer, refreshes the slicer, and saves the file as SlicerShowZeroItemsDemo.xlsx.
class ConfigureSlicerShowZeroItems
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data for the pivot table
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Value");
        cells["A2"].PutValue("A");
        cells["B2"].PutValue(10);
        cells["A3"].PutValue("B");
        cells["B3"].PutValue(20);
        cells["A4"].PutValue("C");
        cells["B4"].PutValue(30);
        // Add a row with no data for category "D"
        cells["A5"].PutValue("D");
        cells["B5"].PutValue(0); // Zero value to represent no data

        // Create a pivot table based on the data range
        int pivotIdx = sheet.PivotTables.Add("A1:B5", "E1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Data, "Value");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a slicer linked to the pivot table for the "Category" field
        int slicerIdx = sheet.Slicers.Add(pivot, "G1", "Category");
        Slicer slicer = sheet.Slicers[slicerIdx];

        // Ensure the slicer shows all items, even those with no data
        slicer.ShowAllItems = true;

        // Configure how items with no data are displayed (e.g., natural order)
        slicer.ShowTypeOfItemsWithNoData = ItemsWithNoDataShowMode.Natural;

        // Optional: refresh the slicer to apply changes
        slicer.Refresh();

        // Save the workbook
        workbook.Save("SlicerShowZeroItemsDemo.xlsx");
    }
}
