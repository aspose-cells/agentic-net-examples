// Title: Select Excel slicer items from a CSV file and refresh the slicer with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells select slicer items based on values from a CSV file | How to programmatically set slicer selections using SlicerCacheItems in Aspose.Cells | Refresh an Excel slicer after changing selection state with Aspose.Cells .NET | Load external CSV data to control Excel slicer selections in C# | Aspose.Cells update slicer cache items from a list of values
// Tags: Aspose.Cells set slicer selections from external data | C# update slicer cache items based on CSV | refresh Excel slicer after programmatic selection | use SlicerCacheItems to control slicer in Aspose.Cells | load CSV values into slicer cache Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;   // Slicer related classes

// The example loads an existing workbook, reads a CSV file containing slicer values, iterates through the slicer's SlicerCacheItems, sets each item's Selected property according to the CSV list, and saves the workbook, effectively updating and refreshing the slicer based on external data.
class SlicerSelectionExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string csvPath = "data.csv";
            const string outputPath = "output.xlsx";

            // Verify input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Workbook file not found: {inputPath}");

            // Load the workbook that contains the slicer
            Workbook workbook = new Workbook(inputPath);

            // Verify CSV file exists
            if (!File.Exists(csvPath))
                throw new FileNotFoundException($"CSV file not found: {csvPath}");

            // Assume the slicer is on the first worksheet and is the first slicer
            Worksheet sheet = workbook.Worksheets[0];
            if (sheet.Slicers.Count == 0)
                throw new InvalidOperationException("No slicers found on the first worksheet.");

            Slicer slicer = sheet.Slicers[0];

            // Get the slicer cache associated with the slicer
            SlicerCache slicerCache = slicer.SlicerCache;

            // Load external CSV data (one column of values to be selected)
            // CSV format: each line contains a value that should be selected in the slicer
            HashSet<string> valuesToSelect = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string line in File.ReadAllLines(csvPath))
            {
                string trimmed = line.Trim();
                if (!string.IsNullOrEmpty(trimmed))
                    valuesToSelect.Add(trimmed);
            }

            // Iterate through all slicer cache items and set their selection state
            foreach (SlicerCacheItem item in slicerCache.SlicerCacheItems)
            {
                // The item value is stored as a string
                string itemValue = item.Value?.ToString() ?? string.Empty;

                // Select the item if its value exists in the CSV list; otherwise deselect it
                item.Selected = valuesToSelect.Contains(itemValue);
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
