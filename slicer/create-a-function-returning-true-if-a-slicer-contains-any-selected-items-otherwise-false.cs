// Title: Determine whether an Excel slicer has any selected items with Aspose.Cells for .NET (C#)
// Common Searches: how to find out if a slicer item is selected using Aspose.Cells in C# | c# Aspose.Cells check for selected slicer values in an Excel file | detect selected items in an Excel slicer with Aspose.Cells .NET API | using Aspose.Cells to verify slicer selection state in a workbook | C# code to test if an Excel slicer contains any checked items via Aspose.Cells
// Tags: aspose.cells slicer selection detection c# | excel slicer item isselected property aspose.cells | c# retrieve sliceritems collection via reflection | aspose.cells workbook slicer inspection | c# check slicer selection status in excel

using System;
using System.IO;
using System.Linq;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Slicers; // Provides Slicer, SlicerItem classes

// The HasSelectedItems method loads a workbook, locates the target worksheet and slicer, uses reflection to access the SlicerItems collection, iterates through each item checking its IsSelected flag, and returns true when any item is selected. It safely handles missing files, absent worksheets or slicers, unavailable properties, and runtime errors by returning false.
public static class SlicerHelper
{
    /// <param name="workbookPath">Full path to the Excel file.</param>
    /// <param name="sheetName">Name of the worksheet that contains the slicer.</param>
    /// <param name="slicerName">Name of the slicer to inspect.</param>
    /// <returns>True if at least one item is selected; false otherwise.</returns>
    public static bool HasSelectedItems(string workbookPath, string sheetName, string slicerName)
    {
        try
        {
            // Prevent FileNotFoundException
            if (!File.Exists(workbookPath))
                return false;

            // Load the workbook
            var workbook = new Workbook(workbookPath);

            // Locate the worksheet (case‑insensitive)
            Worksheet? worksheet = workbook.Worksheets.FirstOrDefault(ws =>
                ws.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase));

            if (worksheet == null)
                return false;

            // Locate the slicer (case‑insensitive)
            Slicer? slicer = worksheet.Slicers.FirstOrDefault(s =>
                s.Name.Equals(slicerName, StringComparison.OrdinalIgnoreCase));

            if (slicer == null)
                return false;

            // Use reflection to obtain the SlicerItems collection (covers API variations)
            var slicerItemsProp = slicer.GetType().GetProperty("SlicerItems");
            if (slicerItemsProp == null)
                return false; // Property not available in this version

            var slicerItemsObj = slicerItemsProp.GetValue(slicer) as IEnumerable;
            if (slicerItemsObj == null)
                return false;

            // Iterate through items and check selection state
            foreach (var itemObj in slicerItemsObj)
            {
                var isSelectedProp = itemObj.GetType().GetProperty("IsSelected");
                if (isSelectedProp != null && isSelectedProp.PropertyType == typeof(bool))
                {
                    bool isSelected = (bool)isSelectedProp.GetValue(itemObj);
                    if (isSelected)
                        return true;
                }
            }

            return false;
        }
        catch (Exception)
        {
            // In case of unexpected errors, treat as no selection
            return false;
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Example usage – adjust the path, sheet, and slicer names as needed.
        string workbookPath = "sample.xlsx";
        string sheetName = "Sheet1";
        string slicerName = "Slicer1";

        bool hasSelection = SlicerHelper.HasSelectedItems(workbookPath, sheetName, slicerName);
        Console.WriteLine($"Slicer '{slicerName}' has selected items: {hasSelection}");
    }
}
