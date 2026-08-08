// Title: Aspose.Cells C# – Show Zero‑Value Items in a Pivot Slicer (ShowAllItems & ItemsWithNoDataShowMode)
// Description: Demonstrates how to create a workbook, add a pivot table, attach a slicer to the "Category" field, enable ShowAllItems, set ItemsWithNoDataShowMode to Natural, refresh the slicer, and save the file so that categories without data appear in the slicer.
// Keywords: Aspose.Cells slicer ShowAllItems | ItemsWithNoDataShowMode Natural | C# pivot table slicer zero items | display empty categories Aspose.Cells | .NET Excel slicer configuration | show zero‑value items in slicer | Aspose.Cells Slicer API | pivot slicer show all items | Excel slicer no data items
// Common Searches: Aspose.Cells show items with no data in slicer | C# slicer ShowAllItems property example | How to display empty categories in Excel slicer using Aspose | ItemsWithNoDataShowMode Natural usage | Pivot table slicer zero‑value items .NET
// Developer Intent: Configure a slicer to list categories even when they have no associated data values.
// Use Cases: Build dashboards where users can filter on categories that may have zero records. | Generate reports that must preserve the full list of items regardless of data presence. | Standardize slicer behavior across multiple workbooks by programmatically enabling ShowAllItems and defining the no‑data display mode.
// AI Prompts: Write C# code with Aspose.Cells to enable a slicer to show categories that have no data and keep natural ordering. | Explain the impact of ShowAllItems and ItemsWithNoDataShowMode on slicer output in Aspose.Cells. | Provide a complete example that creates a pivot table, adds a slicer, configures zero‑value item display, and saves the workbook.

using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// Demonstrates how to create a workbook, add a pivot table, attach a slicer to the "Category" field, enable ShowAllItems, set ItemsWithNoDataShowMode to Natural, refresh the slicer, and save the file so that categories without data appear in the slicer.
class ConfigureSlicerShowZeroItems
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate worksheet with sample data (note that B3 is left empty)
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("B"); // No corresponding value in B3
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B4"].PutValue(30);

        // Add a pivot table based on the data range
        int pivotIndex = worksheet.PivotTables.Add("A1:B4", "E1", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Add a slicer linked to the pivot table for the "Category" field
        int slicerIndex = worksheet.Slicers.Add(pivotTable, "G1", "Category");
        Slicer slicer = worksheet.Slicers[slicerIndex];

        // Ensure the slicer shows all items, even those without data
        slicer.ShowAllItems = true;

        // Configure how items with no data are displayed (e.g., natural order)
        slicer.ShowTypeOfItemsWithNoData = ItemsWithNoDataShowMode.Natural;

        // Refresh the slicer to apply changes
        slicer.Refresh();

        // Save the workbook
        workbook.Save("SlicerShowZeroItems.xlsx");
    }
}
