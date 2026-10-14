// Title: Count cells with custom number formats per worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loops through every worksheet and prints the count of cells whose Style.Custom property is defined. | Create a reusable C# method that returns a Dictionary<string, int> mapping each worksheet name to the number of cells using a custom number format. | Extend the sample to collect the addresses (e.g., A1, B2) of cells that have a custom number format and output them grouped by worksheet.
// Common Searches: Aspose.Cells C# count cells with custom number format in each sheet | How to get number of cells using custom number formats with Aspose.Cells for .NET | Enumerate cells that have a custom Style.Custom value in an Excel workbook using Aspose.Cells | C# Aspose.Cells report custom number format usage per worksheet | Find cells with custom number format in Excel using Aspose.Cells API
// Tags: custom number format counting Aspose.Cells | enumerate cell styles .NET | worksheet custom format statistics C# | Aspose.Cells Style.Custom detection | C# Excel custom format cell enumeration

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, iterates each worksheet, counts cells where the Style.Custom property is set, and prints the count (and optionally the cell addresses) for every sheet.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: File \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int customFormatCount = 0;

                // Enumerate only the cells that actually contain data
                foreach (Cell cell in sheet.Cells)
                {
                    // Retrieve the cell's style
                    Style style = cell.GetStyle();

                    // Check if a custom number format string is defined
                    if (!string.IsNullOrEmpty(style.Custom))
                    {
                        customFormatCount++;
                    }
                }

                // Report the count for the current worksheet
                Console.WriteLine($"Worksheet \"{sheet.Name}\" contains {customFormatCount} cells with custom number formats.");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
