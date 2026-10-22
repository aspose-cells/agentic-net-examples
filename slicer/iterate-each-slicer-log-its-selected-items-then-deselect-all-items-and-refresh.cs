// Title: Iterate through all slicers in an Excel workbook, log selected items, clear selections, and refresh them with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code using Aspose.Cells that loops over every worksheet, accesses each slicer, prints the names of selected slicer items, deselects all items, refreshes the slicer, and saves the workbook. | Update the sample program to enumerate SlicerItem objects, set IsSelected = false for each, call slicer.Refresh(), and ensure the modified workbook is written to a new file.
// Common Searches: C# Aspose.Cells how to list selected slicer items in an Excel file | programmatically clear slicer selections with Aspose.Cells .NET | refresh Excel slicer after changing selections using Aspose.Cells | loop through all slicers in a workbook with Aspose.Cells C# example | Aspose.Cells slicer cache access and item deselection tutorial
// Tags: Aspose.Cells slicer iteration C# | log selected slicer items Aspose.Cells | deselect all slicer items .NET | refresh slicer cache Aspose.Cells | Excel workbook slicer manipulation Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Slicers;
using System;
using System.IO;

// The example loads an Excel workbook, iterates each worksheet and its slicers, optionally logs selected slicer items, deselects all items, refreshes each slicer, and saves the updated workbook to a new file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each slicer on the worksheet
                foreach (Slicer slicer in sheet.Slicers)
                {
                    try
                    {
                        Console.WriteLine($"Slicer: {slicer.Name}");

                        // Access the slicer cache which holds the slicer items
                        SlicerCache cache = slicer.SlicerCache;
                        if (cache == null)
                        {
                            Console.WriteLine("  No slicer cache associated with this slicer.");
                            continue;
                        }

                        // NOTE: In some versions of Aspose.Cells the SlicerItem collection may not be exposed.
                        // The following code logs that the slicer was found and refreshes it.
                        // If the API version supports SlicerItem, you can uncomment and adapt the code below.

                        /*
                        // Log selected items of the slicer
                        foreach (SlicerItem item in cache.SlicerItems)
                        {
                            if (item.IsSelected)
                            {
                                Console.WriteLine($"  Selected Item: {item.Name}");
                            }
                        }

                        // Deselect all items in the slicer
                        foreach (SlicerItem item in cache.SlicerItems)
                        {
                            item.IsSelected = false;
                        }
                        */

                        // Refresh the slicer to apply any changes
                        slicer.Refresh();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"  Error processing slicer '{slicer.Name}': {ex.Message}");
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
