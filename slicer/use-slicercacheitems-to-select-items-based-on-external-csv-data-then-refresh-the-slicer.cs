// Title: Select Excel slicer items from CSV data and refresh slicers using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# program that loads an Excel workbook with Aspose.Cells, reads values from a CSV file, marks matching SlicerCacheItem objects as selected, refreshes each slicer, and saves the updated workbook. | Generate .NET code that iterates through all worksheets, accesses every Slicer, updates its SlicerCache based on external CSV values, calls Refresh on the slicer, and writes the result to a new file.
// Common Searches: C# Aspose.Cells how to programmatically select slicer items using values from a CSV file | update Excel slicer selections from external data source with Aspose.Cells .NET | refresh slicer after changing SlicerCacheItem.Selected property in Aspose.Cells | iterate all worksheets and slicers to apply CSV-based selections in C# | Aspose.Cells example for syncing slicer selections with a CSV list
// Tags: Aspose.Cells slicer cache item selection | C# update slicer selections from CSV | Aspose.Cells refresh slicer programmatically | Excel slicer selection using external data | iterate worksheets slicers Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Slicers; // Required for slicer related classes

// The example loads an existing workbook, reads distinct values from a CSV file, loops through every worksheet and its slicers, sets each SlicerCacheItem.Selected flag according to the CSV list, refreshes the slicers to apply the changes, and saves the modified workbook as a new file.
class SlicerSelectionFromCsv
{
    static void Main()
    {
        try
        {
            // Verify input workbook exists
            const string inputWorkbookPath = "InputWorkbook.xlsx";
            if (!File.Exists(inputWorkbookPath))
                throw new FileNotFoundException($"Workbook file not found: {inputWorkbookPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputWorkbookPath);

            // Verify CSV file exists
            const string csvPath = "SelectionValues.csv";
            if (!File.Exists(csvPath))
                throw new FileNotFoundException($"CSV file not found: {csvPath}");

            // Load external CSV data (one value per line)
            HashSet<string> csvValues = new HashSet<string>(
                File.ReadAllLines(csvPath)
                    .Select(line => line.Trim())
                    .Where(line => !string.IsNullOrEmpty(line)),
                StringComparer.OrdinalIgnoreCase);

            // Iterate through all worksheets to find slicers
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Each worksheet may contain multiple slicers
                foreach (Slicer slicer in sheet.Slicers)
                {
                    // Access the slicer cache associated with this slicer
                    SlicerCache cache = slicer.SlicerCache;

                    // Loop through all items in the slicer cache
                    foreach (SlicerCacheItem item in cache.SlicerCacheItems)
                    {
                        // Select items that exist in the CSV list; otherwise deselect
                        item.Selected = csvValues.Contains(item.Value);
                    }

                    // Refresh the slicer to apply the new selections
                    slicer.Refresh();
                }
            }

            // Save the modified workbook
            const string outputWorkbookPath = "OutputWorkbook.xlsx";
            workbook.Save(outputWorkbookPath);
            Console.WriteLine($"Workbook saved successfully to '{outputWorkbookPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
