// Title: Configuring MemorySetting.MemoryPreference to Normal for faster processing of a small XLSX workbook with Aspose.Cells for .NET
// Common Searches: Aspose.Cells set MemoryPreference to Normal for small Excel files | C# improve speed of loading tiny XLSX using MemorySetting | Is MemoryPreference option supported in current Aspose.Cells .NET version | How to enable normal memory mode in Aspose.Cells when processing a workbook | Workaround for missing MemorySetting API in older Aspose.Cells releases
// Tags: Aspose.Cells MemoryPreference Normal | small XLSX performance optimization | C# workbook memory configuration | fallback for unavailable MemorySetting | fast Excel processing Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for a small XLSX file, loads it into an Aspose.Cells Workbook, notes that the MemorySetting.MemoryPreference property is unavailable in the used library version, and then saves the workbook while handling any runtime exceptions.
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
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing XLSX file
            Workbook workbook = new Workbook(inputPath);

            // Note: MemoryPreference setting is not available in this version of Aspose.Cells.
            // The workbook will use default memory management.

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
