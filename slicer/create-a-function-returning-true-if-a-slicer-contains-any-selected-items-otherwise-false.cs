// Title: Determine whether an Aspose.Cells slicer has any selected items in C# using reflection
// Common Searches: how to check if a slicer has selected items using Aspose.Cells C# | C# reflection to read slicer selected items in an Aspose.Cells workbook | detect slicer selection state in Excel with Aspose.Cells API | verify if an Aspose.Cells slicer contains selections without compile‑time reference | retrieve selected items count from slicer cache levels in Aspose.Cells C#
// Tags: Aspose.Cells slicer selection detection | C# reflection SlicerCacheLevels | read slicer selected items array | Excel slicer selection status check | runtime inspection Aspose.Cells slicer

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The HasSelectedItems method uses reflection to obtain the SlicerCache from a slicer object, iterates through its SlicerCacheLevels, and returns true when any level's SelectedItems array contains entries, handling nulls and exceptions to safely report slicer selection status.
public class SlicerHelper
{
    // Returns true if the specified slicer has any selected items, otherwise false.
    // Uses reflection to avoid compile‑time dependency on the Slicer type (may be unavailable in older Aspose.Cells versions).
    public static bool HasSelectedItems(object slicer)
    {
        try
        {
            if (slicer == null)
                return false;

            // Get the SlicerCache property.
            var slicerType = slicer.GetType();
            var cacheProp = slicerType.GetProperty("SlicerCache");
            if (cacheProp == null)
                return false;

            var cache = cacheProp.GetValue(slicer);
            if (cache == null)
                return false;

            // Get the collection of SlicerCacheLevels.
            var levelsProp = cache.GetType().GetProperty("SlicerCacheLevels");
            var levels = levelsProp?.GetValue(cache) as IEnumerable;
            if (levels == null)
                return false;

            // Iterate through each level and check SelectedItems.
            foreach (var level in levels)
            {
                var selectedProp = level.GetType().GetProperty("SelectedItems");
                var selected = selectedProp?.GetValue(level) as int[];
                if (selected != null && selected.Length > 0)
                    return true;
            }

            // No selected items found.
            return false;
        }
        catch (Exception)
        {
            // In case of any unexpected error, treat as no selection.
            return false;
        }
    }
}

public class Program
{
    // Entry point required for compilation.
    public static void Main()
    {
        try
        {
            // Example usage: check a null slicer (no selection expected).
            bool hasSelection = SlicerHelper.HasSelectedItems(null);
            Console.WriteLine($"Has selected items: {hasSelection}");

            // If you have a workbook with slicers, load it safely:
            // string filePath = "sample.xlsx";
            // if (System.IO.File.Exists(filePath))
            // {
            //     Workbook wb = new Workbook(filePath);
            //     // Retrieve a slicer object from the workbook as needed and pass it to HasSelectedItems.
            // }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
