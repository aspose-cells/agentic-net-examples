// Title: How to programmatically select specific slicer items in an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an existing .xlsx file using Aspose.Cells, iterates over every slicer on a worksheet, and marks as selected the items whose names appear in a predefined list. | Demonstrate how to obtain the SlicerCache for each slicer, use reflection to access its SlicerItems collection, and set the IsSelected property for matching entries. | Provide a complete example that saves the workbook after updating slicer selections, includes error handling for missing files, and creates the output directory if needed.
// Common Searches: Aspose.Cells C# set specific slicer values programmatically | how to mark slicer items as selected from a list using Aspose.Cells | C# iterate over worksheet slicers with Aspose.Cells | use reflection to modify slicer items in an Excel file via Aspose | update multiple slicer selections in an existing workbook .NET
// Tags: Aspose.Cells slicer value selection C# | Excel slicer cache manipulation Aspose | set slicer IsSelected property .NET | programmatic slicer filtering Aspose.Cells | load and save workbook with updated slicers | reflection access SlicerItems Aspose

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// The example loads an existing Excel workbook, defines a list of slicer item names to be selected, loops through all slicers on the first worksheet, uses reflection to retrieve each slicer's items, sets the IsSelected flag for items that match the list, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputFile = "input.xlsx";
            string outputFile = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file '{inputFile}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputFile);

            // Values that should be selected in slicers
            List<string> valuesToSelect = new List<string> { "East", "West", "North" };

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Get all slicers on the worksheet
            SlicerCollection slicers = sheet.Slicers;

            // Iterate through each slicer
            foreach (Slicer slicer in slicers)
            {
                try
                {
                    // Access the slicer cache which holds the items
                    SlicerCache cache = slicer.SlicerCache;
                    if (cache == null) continue;

                    // Use reflection to obtain the slicer items collection (avoids version‑specific API issues)
                    var itemsProp = cache.GetType().GetProperty("SlicerItems");
                    var itemsObj = itemsProp?.GetValue(cache) as System.Collections.IEnumerable;
                    if (itemsObj == null) continue;

                    // Loop through all items of the slicer
                    foreach (var item in itemsObj)
                    {
                        var nameProp = item.GetType().GetProperty("Name");
                        var selectProp = item.GetType().GetProperty("IsSelected");
                        if (nameProp == null || selectProp == null || !selectProp.CanWrite) continue;

                        string name = nameProp.GetValue(item) as string;
                        bool shouldSelect = valuesToSelect.Contains(name);
                        selectProp.SetValue(item, shouldSelect);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing slicer '{slicer.Name}': {ex.Message}");
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputFile);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
