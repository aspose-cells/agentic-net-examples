// Title: Load an XLSX workbook without chart objects using Aspose.Cells LoadOptions and verify empty chart collections in C#
// Common Searches: how to prevent charts from loading when opening an Excel file with Aspose.Cells C# | Aspose.Cells LoadOptions filter to exclude chart objects from workbook | check if worksheets contain charts after loading workbook in C# | C# code to load Excel workbook without charts using Aspose.Cells | verify chart collection count for each worksheet after workbook load Aspose.Cells
// Tags: load workbook excluding charts Aspose.Cells | filter objects loadoptions C# | worksheet chart count verification Aspose.Cells | skip chart objects on workbook load | chart collection empty check C#

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading an XLSX file with Aspose.Cells while skipping chart objects via LoadOptions, then iterates through each worksheet to report whether its Charts collection is empty or contains charts, handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook without loading unnecessary objects
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Verify that each worksheet's chart collection is empty (or report count)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Charts.Count == 0)
                {
                    Console.WriteLine($"Worksheet \"{sheet.Name}\" has no charts.");
                }
                else
                {
                    Console.WriteLine($"Worksheet \"{sheet.Name}\" contains {sheet.Charts.Count} chart(s).");
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during loading or processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
