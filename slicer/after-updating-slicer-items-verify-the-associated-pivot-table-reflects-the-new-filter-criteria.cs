// Title: Check that a pivot table reflects updated slicer selections using Aspose.Cells for .NET (C#)
// AI Prompts: Select the first two items of a slicer, deselect the remaining items, refresh the linked pivot table, and output a verification message with Aspose.Cells in C#. | Count the selected slicer cache items and compare the count to the number of visible row items in the pivot table to confirm filter synchronization. | Adapt the code to locate a slicer by its name, apply custom selections, refresh the associated pivot table, and validate the result.
// Common Searches: Aspose.Cells C# how to programmatically change slicer items and refresh linked pivot table | verify slicer filter updates pivot table rows using Aspose.Cells .NET | count visible pivot items after applying slicer selections with Aspose.Cells | C# example for refreshing pivot table after slicer cache item selection Aspose.Cells | Aspose.Cells verify that slicer selections match pivot row items
// Tags: update slicer cache items Aspose.Cells | refresh pivot table after slicer change .NET | verify slicer filter synchronization with pivot table | count visible pivot items C# Aspose.Cells | select slicer items programmatically Excel .NET | pivot table data refresh Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

// The example loads an Excel workbook, accesses the first slicer on the first worksheet, selects the first two slicer items while deselecting the rest, refreshes the first pivot table on the sheet, counts both the selected slicer items and the visible row items in the pivot table, compares the counts to verify that the pivot reflects the slicer filter, outputs the verification result, and saves the workbook.
class SlicerPivotVerification
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook containing the slicer and pivot table
            Workbook workbook = new Workbook(inputPath);

            // Assume the slicer is on the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one slicer
            if (sheet.Slicers.Count == 0)
            {
                Console.WriteLine("No slicers found on the first worksheet.");
                return;
            }

            // Access the first slicer
            Slicer slicer = sheet.Slicers[0];

            // Update slicer items: select the first two items and deselect the rest
            SlicerCache slicerCache = slicer.SlicerCache;
            var cacheItems = slicerCache.SlicerCacheItems;
            for (int i = 0; i < cacheItems.Count; i++)
            {
                cacheItems[i].Selected = i < 2;
            }

            // Retrieve a linked pivot table (use the first pivot table on the sheet)
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
                return;
            }

            PivotTable pivotTable = sheet.PivotTables[0];

            // Refresh the pivot table to apply the slicer filter
            pivotTable.RefreshData();

            // Count selected slicer items
            int selectedSlicerCount = 0;
            foreach (SlicerCacheItem item in cacheItems)
            {
                if (item.Selected)
                    selectedSlicerCount++;
            }

            // Ensure the pivot table has at least one row field
            if (pivotTable.RowFields.Count == 0)
            {
                Console.WriteLine("Pivot table has no row fields to verify.");
                return;
            }

            // Get the first row field of the pivot table
            PivotField rowField = pivotTable.RowFields[0];

            // Count visible items in the row field after filtering
            int visibleRowItemCount = 0;
            foreach (PivotItem pivotItem in rowField.PivotItems)
            {
                if (!pivotItem.IsHidden) // IsHidden = false means the item is visible after filter
                    visibleRowItemCount++;
            }

            // Output verification result
            if (selectedSlicerCount == visibleRowItemCount)
            {
                Console.WriteLine("Verification succeeded: Pivot table reflects slicer filter.");
            }
            else
            {
                Console.WriteLine($"Verification failed: Selected slicer items = {selectedSlicerCount}, " +
                                  $"visible pivot items = {visibleRowItemCount}");
            }

            // Save the workbook if needed
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
