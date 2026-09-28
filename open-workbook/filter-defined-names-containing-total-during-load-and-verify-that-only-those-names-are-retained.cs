// Title: Load an Excel workbook using Aspose.Cells for .NET and retain only defined names containing the word "Total"
// AI Prompts: Load a workbook with Aspose.Cells, iterate through the NameCollection, delete every defined name whose Text does not include the substring "Total" (case‑insensitive), verify the remaining names, and save the file. | Write C# code that filters the defined names in an Excel file so that after loading, only names containing "Total" remain, includes error handling for a missing file and throws an exception if any retained name lacks the keyword.
// Common Searches: Aspose.Cells C# filter workbook defined names containing specific keyword | How to remove named ranges that do not include 'Total' using Aspose.Cells | Retain only Excel defined names with 'Total' after loading workbook in .NET | Verify defined name substrings in Aspose.Cells after opening a file | C# Aspose.Cells remove unwanted named ranges based on text
// Tags: Aspose.Cells defined name filtering | C# remove Excel named ranges by keyword | retain named ranges containing Total | NameCollection manipulation Aspose.Cells | verify defined name substring .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Loads Input.xlsx with Aspose.Cells, removes every defined name whose Text does not contain the word "Total" (case‑insensitive), validates that all remaining names include the keyword, and saves the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Access the collection of defined names in the workbook
            NameCollection definedNames = workbook.Worksheets.Names;

            // Collect the names that do NOT contain "Total"
            List<string> namesToRemove = new List<string>();
            foreach (Name definedName in definedNames)
            {
                // Use the Text property to get the defined name string
                if (!definedName.Text.Contains("Total", StringComparison.OrdinalIgnoreCase))
                {
                    namesToRemove.Add(definedName.Text);
                }
            }

            // Remove the unwanted defined names from the collection
            foreach (string name in namesToRemove)
            {
                definedNames.Remove(name);
            }

            // Verification: ensure every remaining defined name contains "Total"
            foreach (Name definedName in definedNames)
            {
                if (!definedName.Text.Contains("Total", StringComparison.OrdinalIgnoreCase))
                {
                    throw new Exception($"Defined name '{definedName.Text}' does not contain 'Total'.");
                }
            }

            // Save the workbook (optional, to persist the filtered defined names)
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
