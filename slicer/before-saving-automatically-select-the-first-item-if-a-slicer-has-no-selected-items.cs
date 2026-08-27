// Title: How to automatically select the first slicer item in an Excel workbook using Aspose.Cells for .NET before saving
// AI Prompts: Generate C# code with Aspose.Cells that loads an Excel file, iterates every slicer, and if a slicer has no selected items, programmatically marks the first cache entry as selected and calls Refresh before saving. | Update an existing C# Aspose.Cells routine to guarantee a default selection for each slicer by checking the SlicerCacheItems collection, setting the first item's Selected property when none are true, and then persisting the workbook.
// Common Searches: how to ensure a slicer always has a selected value with Aspose.Cells .NET | c# aspocells default slicer selection before workbook save | iterate over slicers in Excel using Aspose.Cells and set first item selected | refresh slicer after changing selection programmatically aspocells | check slicer cache items for selection aspocells c#
// Tags: aspocells default slicer selection | c# iterate workbook slicers | set first slicer cache item selected | slicer refresh after programmatic change | excel workbook slicer handling .net

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

// The program loads an existing workbook, refreshes all data, loops through each worksheet and its slicers, checks whether any slicer cache items are selected, selects the first item when none are selected, refreshes the slicer, and saves the workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook that contains slicers
        Workbook workbook = new Workbook("input.xlsx");

        // Ensure all data (pivot tables, slicer caches) are up‑to‑date
        workbook.Worksheets.RefreshAll();

        // Loop through every worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Loop through each slicer on the current worksheet
            foreach (Slicer slicer in sheet.Slicers)
            {
                // Access the collection of slicer cache items
                var cacheItems = slicer.SlicerCache.SlicerCacheItems;

                // Determine whether any item is already selected
                bool anySelected = false;
                for (int i = 0; i < cacheItems.Count; i++)
                {
                    if (cacheItems[i].Selected)
                    {
                        anySelected = true;
                        break;
                    }
                }

                // If no items are selected, select the first item
                if (!anySelected && cacheItems.Count > 0)
                {
                    cacheItems[0].Selected = true;
                    // Optionally set the slicer's first visible item index to 0
                    slicer.FirstItemIndex = 0;
                }

                // Refresh the slicer so the selection takes effect
                slicer.Refresh();
            }
        }

        // Save the workbook after the automatic selection has been applied
        workbook.Save("output.xlsx");
    }
}
