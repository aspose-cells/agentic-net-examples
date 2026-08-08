// Title: Aspose.Cells for .NET – Enable Multi‑Select on a Pivot Table Slicer (C#)
// Description: This C# sample builds a workbook with fruit sales data, creates a pivot table, activates multiple‑item selection on the row field by setting the page field’s IsMultipleItemSelectionAllowed property, adds a slicer linked to the Fruit field, programmatically selects the first two items, refreshes the slicer, and saves the file as SlicerMultiSelectDemo.xlsx.
// Keywords: Aspose.Cells | C# | pivot table slicer | multi‑select slicer | IsMultipleItemSelectionAllowed | programmatic slicer selection | Excel slicer .NET | Aspose.Cells API | slicer cache items
// Common Searches: Aspose.Cells enable slicer multi select | C# set slicer to allow multiple items | How to programmatically select slicer items Aspose.Cells | Multi‑select pivot slicer Aspose.Cells .NET | Set IsMultipleItemSelectionAllowed property
// Developer Intent: Configure a slicer so users can pick several items at once and set default selections via code.
// Use Cases: Allow end‑users to filter a pivot table by multiple fruit categories simultaneously. | Pre‑select Apple and Orange when the workbook opens to display combined sales. | Drive charts and tables in a dashboard where slicer multi‑select controls the data view. | Automate report generation that requires specific slicer filters applied programmatically.
// AI Prompts: Generate C# code that enables multi‑select on an Aspose.Cells slicer linked to a pivot table. | Show how to set IsMultipleItemSelectionAllowed and pre‑select slicer items using Aspose.Cells for .NET. | Explain the steps to add a slicer, enable multiple selections, and refresh it in a workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerMultiSelectDemo
{
    // This C# sample builds a workbook with fruit sales data, creates a pivot table, activates multiple‑item selection on the row field by setting the page field’s IsMultipleItemSelectionAllowed property, adds a slicer linked to the Fruit field, programmatically selects the first two items, refreshes the slicer, and saves the file as SlicerMultiSelectDemo.xlsx.
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
            cells["B1"].Value = "Sales";
            cells["A2"].Value = "Apple";
            cells["B2"].Value = 120;
            cells["A3"].Value = "Orange";
            cells["B3"].Value = 150;
            cells["A4"].Value = "Banana";
            cells["B4"].Value = 90;
            cells["A5"].Value = "Apple";
            cells["B5"].Value = 80;
            cells["A6"].Value = "Orange";
            cells["B6"].Value = 70;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B6", "D3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            // Row field (the slicer will be based on this field)
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            // Data field
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.RefreshData();
            pivot.CalculateData();

            // Enable multiple item selection for the page field (required for slicer multi‑select)
            // The row field automatically becomes a page field when a slicer is added.
            // Access the first page field and set the property.
            if (pivot.PageFields.Count > 0)
            {
                PivotField pageField = pivot.PageFields[0];
                pageField.IsMultipleItemSelectionAllowed = true;
            }

            // Add a slicer linked to the "Fruit" field of the pivot table
            int slicerIndex = sheet.Slicers.Add(pivot, "F3", "Fruit");
            Slicer slicer = sheet.Slicers[slicerIndex];
            slicer.StyleType = SlicerStyleType.SlicerStyleLight1;
            slicer.Caption = "Fruit Slicer";

            // Select multiple items in the slicer (e.g., first two items)
            for (int i = 0; i < slicer.SlicerCache.SlicerCacheItems.Count; i++)
            {
                SlicerCacheItem item = slicer.SlicerCache.SlicerCacheItems[i];
                // Select the first two items, deselect the rest
                item.Selected = i < 2;
            }

            // Refresh the slicer to apply the selection changes
            slicer.Refresh();

            // Save the workbook
            workbook.Save("SlicerMultiSelectDemo.xlsx");
        }
    }
}
