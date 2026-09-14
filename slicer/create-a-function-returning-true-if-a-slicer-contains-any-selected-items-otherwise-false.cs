// Title: C# utility method to check whether an Aspose.Cells slicer contains any selected items
// AI Prompts: Generate a static C# method that accepts an Aspose.Cells Slicer object, validates it, and returns true if any SlicerCacheItem has its Selected flag set. | Write C# code that iterates through slicer.SlicerCache.SlicerCacheItems and determines if at least one item is selected, throwing ArgumentNullException for a null slicer.
// Common Searches: how to verify if a slicer in Aspose.Cells for .NET has selected values | C# Aspose.Cells method to detect selected slicer cache items | check slicer selection status using Aspose.Cells API | Aspose.Cells SlicerHelper HasSelectedItems usage example | determine if any slicer items are selected in an Excel workbook with C#
// Tags: Aspose.Cells slicer selection check | C# detect slicer cache item selected | SlicerCacheItem Selected property Aspose.Cells | validate slicer reference null handling | utility method for slicer selection status | Excel slicer selected items detection C#

using System;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerUtilities
{
    // Provides a static helper method HasSelectedItems that throws ArgumentNullException for a null slicer, iterates through slicer.SlicerCache.SlicerCacheItems, and returns true when any item's Selected property is true; otherwise returns false.
    public static class SlicerHelper
    {
        /// <param name="slicer">The slicer to inspect.</param>
        /// <returns>True if at least one slicer cache item is selected; otherwise false.</returns>
        public static bool HasSelectedItems(Slicer slicer)
        {
            // Guard against null slicer reference
            if (slicer == null)
                throw new ArgumentNullException(nameof(slicer));

            // Access the collection of cache items associated with the slicer
            SlicerCacheItemCollection items = slicer.SlicerCache.SlicerCacheItems;

            // Iterate through the items and check the Selected property
            for (int i = 0; i < items.Count; i++)
            {
                SlicerCacheItem item = items[i];
                if (item.Selected)
                {
                    // At least one item is selected
                    return true;
                }
            }

            // No selected items found
            return false;
        }
    }

    // Entry point for the console application
    internal class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                // Placeholder for future demonstration code.
                // Example: load a workbook, retrieve a slicer, and call HasSelectedItems.
                // Currently no operation is required.
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
