// Title: Check if an Aspose.Cells Slicer Has Selected Items (C#/.NET)
// Description: Provides a C# helper method that inspects a slicer's SlicerCacheItems collection and returns true when at least one item’s Selected flag is set. Includes sample code that loads a workbook, retrieves the first slicer, and prints the selection status, with argument validation and error handling.
// Keywords: Aspose.Cells slicer selected items | C# check slicer selection | SlicerCacheItem.Selected | Excel slicer detection .NET | Aspose.Cells SlicerHelper | boolean slicer check | workbook slicer API | Aspose.Cells .NET example
// Common Searches: how to detect selected items in an Aspose.Cells slicer | C# Aspose.Cells check if slicer has any selections | determine slicer selection state using Aspose.Cells for .NET | Aspose.Cells slicer selected flag example | C# method to verify slicer selections in Excel workbook
// Developer Intent: Find out whether a slicer contains at least one selected cache item.
// Use Cases: Prevent applying filters when a slicer has no selections, avoiding empty result sets. | Toggle UI controls (e.g., buttons, panels) based on slicer activity. | Log user interaction with slicers by recording when selections exist.
// AI Prompts: Generate unit tests for SlicerHelper.HasSelectedItems covering both selected and unselected scenarios using Moq. | Rewrite HasSelectedItems with LINQ to make the implementation a single expression. | Create an overload of HasSelectedItems that returns the count of selected items instead of a boolean.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// Provides a C# helper method that inspects a slicer's SlicerCacheItems collection and returns true when at least one item’s Selected flag is set. Includes sample code that loads a workbook, retrieves the first slicer, and prints the selection status, with argument validation and error handling.
public static class SlicerHelper
{
    // Returns true if the slicer has at least one selected item
    public static bool HasSelectedItems(Slicer slicer)
    {
        if (slicer == null) throw new ArgumentNullException(nameof(slicer));

        // Access the collection of cache items for the slicer
        SlicerCacheItemCollection items = slicer.SlicerCache.SlicerCacheItems;

        // Iterate through items and check the Selected property
        foreach (SlicerCacheItem item in items)
        {
            if (item.Selected)
                return true; // Found a selected item
        }

        return false; // No selected items found
    }
}

class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"File '{inputPath}' not found.");
                return;
            }

            // Load the workbook that contains a slicer
            Workbook workbook = new Workbook(inputPath);
            workbook.Worksheets.RefreshAll();

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("Workbook contains no worksheets.");
                return;
            }

            Worksheet ws = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one slicer
            if (ws.Slicers == null || ws.Slicers.Count == 0)
            {
                Console.WriteLine("No slicers found in the first worksheet.");
                return;
            }

            // Retrieve the first slicer
            Slicer slicer = ws.Slicers[0];

            // Determine if any items are selected
            bool anySelected = SlicerHelper.HasSelectedItems(slicer);
            Console.WriteLine("Slicer has selected items: " + anySelected);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
