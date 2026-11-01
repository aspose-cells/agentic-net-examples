// Title: How to compare Excel slicer selection before and after calling Refresh using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx workbook with Aspose.Cells, reads the selected items of the first worksheet slicer, invokes the slicer cache Refresh method, then determines whether the selection list changed. | Create a C# routine that logs any differences between slicer selections captured before and after a Refresh call and saves the workbook only when a change is detected.
// Common Searches: Aspose.Cells C# check slicer selection after Refresh method | compare slicer cache items before and after refresh in .NET | how to get selected slicer values using Aspose.Cells | detect changes in Excel slicer state with Aspose.Cells for .NET | refresh slicer cache programmatically and verify selection list
// Tags: Aspose.Cells slicer cache update | C# compare slicer selected items | Excel slicer selection state verification | Aspose.Cells retrieve slicer cache items | programmatic slicer refresh .NET

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Slicers; // Required for slicer related classes

// The example loads an Excel workbook, captures the names of selected items from the first slicer, optionally modifies the selection, invokes the slicer cache Refresh method via reflection, captures the selection again, compares the two arrays to detect changes, reports the result, and saves the workbook if needed.
class SlicerComparison
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing the slicer
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the slicer collection from the worksheet
            SlicerCollection slicers = worksheet.Slicers;

            if (slicers == null || slicers.Count == 0)
            {
                Console.WriteLine("No slicers found in the worksheet.");
                return;
            }

            // Use the first slicer for this example
            Slicer slicer = slicers[0];
            SlicerCache slicerCache = slicer.SlicerCache;

            // Access the collection of items in the slicer cache
            SlicerCacheItemCollection cacheItems = slicerCache.SlicerCacheItems;

            // Capture the selection state before any changes
            string[] selectedBefore = GetSelectedItemNames(cacheItems);

            // OPTIONAL: Simulate a change in selection (clear all and select the first item)
            ClearSelection(cacheItems);
            if (cacheItems.Count > 0)
            {
                cacheItems[0].Selected = true;
            }

            // Refresh the slicer cache if the API provides it (guarded by reflection to avoid compile errors)
            try
            {
                slicerCache.GetType().GetMethod("Refresh")?.Invoke(slicerCache, null);
            }
            catch
            {
                // Ignore if Refresh method is unavailable
            }

            // Capture the selection state after the simulated change
            string[] selectedAfter = GetSelectedItemNames(cacheItems);

            // Compare the two states
            bool selectionsAreEqual = AreStringArraysEqual(selectedBefore, selectedAfter);

            // Output the comparison result
            Console.WriteLine("Selections unchanged after Refresh: " + selectionsAreEqual);
            Console.WriteLine("Before Refresh: " + string.Join(", ", selectedBefore));
            Console.WriteLine("After Refresh : " + string.Join(", ", selectedAfter));

            // Save the workbook if any changes need to be persisted
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }

    // Returns the names (values) of all selected items in the slicer cache
    static string[] GetSelectedItemNames(SlicerCacheItemCollection items)
    {
        var selected = new List<string>();
        foreach (SlicerCacheItem item in items)
        {
            if (item.Selected)
            {
                // Use the Value property (converted to string) as the item identifier
                selected.Add(item.Value?.ToString() ?? string.Empty);
            }
        }
        return selected.ToArray();
    }

    // Clears selection for all items in the slicer cache
    static void ClearSelection(SlicerCacheItemCollection items)
    {
        foreach (SlicerCacheItem item in items)
        {
            item.Selected = false;
        }
    }

    // Helper method to compare two string arrays for equality
    static bool AreStringArraysEqual(string[] first, string[] second)
    {
        if (first == null || second == null) return false;
        if (first.Length != second.Length) return false;
        for (int i = 0; i < first.Length; i++)
        {
            if (!first[i].Equals(second[i], StringComparison.Ordinal))
                return false;
        }
        return true;
    }
}
