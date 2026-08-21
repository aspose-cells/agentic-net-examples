// Title: Select Multiple Slicer Items, Refresh Pivot, and Save Workbook with Aspose.Cells for .NET (C#)
// Description: Load an XLSX file, retrieve a slicer, programmatically select specific items (e.g., "Apple" and "Banana"), refresh the slicer to update the linked pivot table, and save the modified workbook using Aspose.Cells for C#.
// Keywords: Aspose.Cells slicer selection | C# programmatic slicer | refresh slicer pivot table | save workbook after slicer update | SlicerCacheItem Aspose | Excel slicer automation | multiple slicer values .NET
// Common Searches: how to select multiple values in an Excel slicer using Aspose.Cells C# | refresh slicer after changing selections Aspose.Cells | save workbook with updated slicer filters .NET | Aspose.Cells programmatic slicer example | C# code to filter pivot table via slicer
// Developer Intent: Programmatically set the selection state of slicer items, apply the filter to the associated pivot table, and persist the changes to a new XLSX file.
// Use Cases: Generate a daily sales report that automatically shows only selected product categories. | Create a batch job that updates slicer filters based on external data before exporting the workbook. | Build an interactive tool where user‑chosen slicer values are applied to a pivot table and saved without manual Excel interaction.
// AI Prompts: Write C# code with Aspose.Cells to select "East" and "West" regions in a slicer, refresh the linked pivot table, and save the workbook. | Show how to loop through SlicerCacheItems, deselect all items, then select a custom list of values and persist the changes to an XLSX file using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// Load an XLSX file, retrieve a slicer, programmatically select specific items (e.g., "Apple" and "Banana"), refresh the slicer to update the linked pivot table, and save the modified workbook using Aspose.Cells for C#.
class SlicerSelectionDemo
{
    static void Main()
    {
        // Load an existing workbook that contains a slicer connected to a pivot table
        Workbook workbook = new Workbook("input.xlsx");

        // Assume the slicer is on the first worksheet; adjust index/name as needed
        Worksheet sheet = workbook.Worksheets[0];
        SlicerCollection slicers = sheet.Slicers;

        // Retrieve the first slicer (or use slicers["SlicerName"] if known)
        if (slicers.Count == 0)
        {
            Console.WriteLine("No slicers found in the worksheet.");
            return;
        }

        Slicer slicer = slicers[0];

        // Select multiple items in the slicer.
        // Example: select items whose value is "Apple" or "Banana".
        foreach (SlicerCacheItem item in slicer.SlicerCache.SlicerCacheItems)
        {
            if (item.Value == "Apple" || item.Value == "Banana")
                item.Selected = true;   // select the item
            else
                item.Selected = false;  // deselect other items
        }

        // Refresh the slicer to apply the selection and recalculate the pivot table
        slicer.Refresh();

        // Save the workbook with the updated slicer selections
        workbook.Save("output.xlsx");
    }
}
