// Title: Read worksheet TabId values from an Excel file with Aspose.Cells for .NET using low‑memory LoadOptions
// Common Searches: how to read worksheet TabId in C# with Aspose.Cells without loading the full workbook | Aspose.Cells load Excel file with MemorySetting.MemoryPreference to get sheet identifiers | C# extract internal sheet IDs (TabId) from .xlsx using Aspose.Cells low memory mode | retrieve Excel worksheet TabId from OpenXML package using Aspose.Cells LoadOptions | list worksheet names and TabId values in .NET without loading all cell data
// Tags: Aspose.Cells MemorySetting low‑memory loading | C# read worksheet TabId | extract sheet identifiers from .xlsx | worksheet enumeration without full workbook load | LoadOptions for efficient Excel parsing

using System;
using System.IO;
using Aspose.Cells;

namespace ExtractTabIdDemo
{
    // The example validates a single file‑path argument, creates LoadOptions with MemorySetting.MemoryPreference to minimize memory usage, loads the workbook, iterates through each worksheet, and prints each sheet's name together with its internal TabId, handling any errors that may occur.
    class Program
    {
        /// <param name="args">Command‑line arguments.</param>
        static void Main(string[] args)
        {
            // Validate input arguments
            if (args == null || args.Length != 1)
            {
                Console.WriteLine("Usage: ExtractTabIdDemo <path-to-excel-file>");
                return;
            }

            string filePath = args[0];

            // Ensure the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File not found – \"{filePath}\"");
                return;
            }

            try
            {
                // Configure load options to minimise memory usage.
                // LoadDataOnly is not available in the current Aspose.Cells version,
                // so we rely on MemorySetting for efficient loading.
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto)
                {
                    MemorySetting = MemorySetting.MemoryPreference
                };

                // Load the workbook using the configured options.
                // Using statement ensures the workbook is disposed properly.
                using (Workbook workbook = new Workbook(filePath, loadOptions))
                {
                    Console.WriteLine(
                        $"Workbook \"{Path.GetFileName(filePath)}\" contains {workbook.Worksheets.Count} worksheet(s):");

                    // Iterate through each worksheet and output its TabId.
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Worksheet.TabId provides the internal identifier for the sheet.
                        int tabId = sheet.TabId;
                        Console.WriteLine($"  Sheet Name: \"{sheet.Name}\", TabId: {tabId}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while extracting TabId information:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
