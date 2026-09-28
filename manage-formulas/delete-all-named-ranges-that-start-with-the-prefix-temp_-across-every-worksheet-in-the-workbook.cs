// Title: Remove all named ranges beginning with 'Temp_' from every worksheet in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, finds every named range whose name starts with 'Temp_', and deletes them from all worksheets. | Create a C# method using Aspose.Cells that iterates through workbook names, filters by a prefix, removes the matching named ranges, and saves the file.
// Common Searches: Aspose.Cells C# delete named ranges that start with a prefix | How to programmatically remove temporary named ranges from an Excel file using .NET | C# iterate workbook names and delete those matching 'Temp_' with Aspose.Cells | Clean up Excel named ranges across all sheets in Aspose.Cells
// Tags: aspocells delete named ranges by prefix | c# remove workbook names Aspose.Cells | filter excel named ranges Aspose.Cells | named range cleanup .NET | global and worksheet level name removal Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, gathers all global and worksheet‑level named ranges whose names start with "Temp_", removes each of them, and saves the updated workbook.
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
            // Load the workbook from the input file
            var workbook = new Workbook(inputPath);

            // Collect all named ranges (global and worksheet‑level) that start with 'Temp_'
            var namesToDelete = new List<string>();
            foreach (Name name in workbook.Worksheets.Names)
            {
                if (name.Text.StartsWith("Temp_", StringComparison.OrdinalIgnoreCase))
                {
                    namesToDelete.Add(name.Text);
                }
            }

            // Remove the collected named ranges
            foreach (string name in namesToDelete)
            {
                workbook.Worksheets.Names.Remove(name);
            }

            // Save the modified workbook to the output file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
