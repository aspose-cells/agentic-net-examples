// Title: How to enable multi‑select on an Aspose.Cells slicer linked to a pivot table in C#
// AI Prompts: Generate C# code that configures a pivot field to allow multiple item selection and links a slicer to it using Aspose.Cells. | Show how to programmatically clear existing selections and pre‑select specific slicer items (e.g., Apple and Orange) with Aspose.Cells for .NET. | Explain the steps to refresh a slicer after modifying its selection state in an Aspose.Cells workbook.
// Common Searches: Aspose.Cells C# enable multi select on slicer for pivot table | set IsMultipleItemSelectionAllowed property Aspose.Cells slicer | preselect items in Aspose.Cells slicer programmatically | refresh slicer after changing selection Aspose.Cells .NET | create slicer linked to pivot field with multiple selection Aspose.Cells
// Tags: Aspose.Cells slicer multiple item selection | C# pivot table slicer linking | IsMultipleItemSelectionAllowed usage | programmatic slicer cache item selection | slicer refresh method Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerMultiSelectDemo
{
    // The example creates a workbook, adds sample data and a pivot table, enables multiple item selection on the 'Fruit' pivot field, inserts a slicer linked to that field, clears all selections, pre‑selects 'Apple' and 'Orange', refreshes the slicer, and saves the file as SlicerMultiSelectDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Fruit";
            cells["B1"].Value = "Quantity";
            cells["A2"].Value = "Apple";
            cells["B2"].Value = 10;
            cells["A3"].Value = "Orange";
            cells["B3"].Value = 15;
            cells["A4"].Value = "Banana";
            cells["B4"].Value = 20;
            cells["A5"].Value = "Apple";
            cells["B5"].Value = 5;
            cells["A6"].Value = "Orange";
            cells["B6"].Value = 8;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B6", "D3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add the "Fruit" field to the Row area (this will become the slicer field)
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            // Add the "Quantity" field to the Data area
            pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");

            // Refresh and calculate the pivot table
            pivot.RefreshData();
            pivot.CalculateData();

            // Enable multiple item selection for the page field (the slicer field)
            // The field appears in PageFields after it is added to the slicer, but we can access it via BaseFields
            // Here we use the first BaseField (index 0) which corresponds to "Fruit"
            PivotField fruitField = pivot.BaseFields[0];
            fruitField.IsMultipleItemSelectionAllowed = true;

            // Add a slicer linked to the "Fruit" field
            // Using the overload that takes the field name
            int slicerIndex = sheet.Slicers.Add(pivot, "E3", "Fruit");
            Slicer slicer = sheet.Slicers[slicerIndex];

            // Optional: pre‑select a couple of items to demonstrate multi‑selection
            // By default all items are selected; here we deselect all and then select two items
            foreach (SlicerCacheItem item in slicer.SlicerCache.SlicerCacheItems)
            {
                item.Selected = false; // clear previous selection
            }
            // Select "Apple" and "Orange"
            foreach (SlicerCacheItem item in slicer.SlicerCache.SlicerCacheItems)
            {
                if (item.Value == "Apple" || item.Value == "Orange")
                {
                    item.Selected = true;
                }
            }

            // Refresh the slicer to apply the selection changes
            slicer.Refresh();

            // Save the workbook
            workbook.Save("SlicerMultiSelectDemo.xlsx");
        }
    }
}
