// Title: Remove temporary defined names (temp_*) from an Excel workbook using Aspose.Cells LoadOptions and a custom LoadFilter (C#)
// Description: Shows how to configure LoadOptions with a custom LoadFilter that loads all worksheet data, open a workbook, iterate its NameCollection, and delete any defined name whose text starts with "temp_" before saving the file.
// Keywords: Aspose.Cells | LoadOptions | LoadFilter | defined names | temporary named ranges | C# | Excel | remove temp_ | NameCollection
// Common Searches: Aspose.Cells remove defined names starting with temp_ | C# LoadFilter exclude temporary named ranges | LoadOptions filter defined names Aspose.Cells | How to delete temp_ named ranges in Excel using Aspose | Custom LoadFilter to clean up workbook
// Developer Intent: Load an Excel workbook and automatically discard any defined names that begin with the prefix "temp_".
// Use Cases: Clean up helper named ranges imported from third‑party systems before further processing. | Prepare a workbook for distribution by stripping internal temporary names that users should not see. | Lower memory usage and improve performance by omitting unnecessary defined names during initial load.
// AI Prompts: Generate a C# example that uses Aspose.Cells LoadOptions with a custom LoadFilter to delete defined names starting with "temp_" after the workbook is loaded. | Refactor the provided code so the temporary defined names are removed inside the LoadFilter instead of iterating the NameCollection later. | Explain how LoadDataFilterOptions and a custom LoadFilter can be combined to selectively load worksheet content and exclude specific defined names in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to configure LoadOptions with a custom LoadFilter that loads all worksheet data, open a workbook, iterate its NameCollection, and delete any defined name whose text starts with "temp_" before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Create LoadOptions and assign a custom LoadFilter
            LoadOptions loadOptions = new LoadOptions
            {
                LoadFilter = new TempNameFilter()
            };

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Remove defined names that start with "temp_"
            NameCollection definedNames = workbook.Worksheets.Names;
            for (int i = definedNames.Count - 1; i >= 0; i--)
            {
                Name definedName = definedNames[i];
                if (definedName.Text.StartsWith("temp_", StringComparison.OrdinalIgnoreCase))
                {
                    definedNames.RemoveAt(i);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Custom LoadFilter that loads all data for each worksheet
    private class TempNameFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Ensure all data (including defined names) are loaded
            this.LoadDataFilterOptions = LoadDataFilterOptions.All;
        }
    }
}
