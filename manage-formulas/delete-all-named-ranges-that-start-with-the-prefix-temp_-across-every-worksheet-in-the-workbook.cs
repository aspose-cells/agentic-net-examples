// Title: How to delete all workbook‑level named ranges that start with the 'Temp_' prefix across every worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# remove all named ranges beginning with Temp_ from a workbook | Delete workbook level names that have a specific prefix in Aspose.Cells .NET | Programmatically filter and delete Excel named ranges with a prefix using Aspose.Cells | C# iterate NameCollection in reverse to remove temporary named ranges | How to clean up temporary named ranges across all sheets with Aspose.Cells
// Tags: remove workbook named ranges Aspose.Cells | delete named ranges with prefix .NET | NameCollection iteration removal Aspose.Cells | clean temporary Excel names C# | Aspose.Cells workbook cleanup temporary names

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, accesses its NameCollection, iterates the collection in reverse, removes any named range whose text starts with "Temp_", and saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the collection of named ranges (workbook‑level)
            NameCollection names = workbook.Worksheets.Names;

            // Remove named ranges that start with "Temp_"
            // Iterate in reverse order to safely remove items while looping
            for (int i = names.Count - 1; i >= 0; i--)
            {
                Name name = names[i];
                if (name.Text.StartsWith("Temp_", StringComparison.OrdinalIgnoreCase))
                {
                    names.RemoveAt(i);
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
}
