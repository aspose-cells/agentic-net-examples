// Title: Remove defined names prefixed with temp_ from an Excel workbook using Aspose.Cells in C#
// Common Searches: how to delete named ranges that start with temp_ using Aspose.Cells C# | Aspose.Cells remove temporary defined names from workbook programmatically | C# code to filter out named ranges with specific prefix in Excel file | exclude defined names with prefix temp_ when processing Excel with Aspose.Cells | remove unwanted named ranges during workbook load Aspose.Cells .NET
// Tags: Aspose.Cells remove defined names C# | filter named ranges by prefix Aspose.Cells | delete temporary named ranges Excel .NET | C# Excel defined name cleanup Aspose | Aspose.Cells workbook name collection removal

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, iterates through the workbook's defined names, identifies those whose text starts with "temp_", removes the matching names from the name collection, and saves the cleaned workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Collect defined names that start with "temp_"
            NameCollection nameCollection = workbook.Worksheets.Names;
            List<string> namesToRemove = new List<string>();

            foreach (Name definedName in nameCollection)
            {
                if (!string.IsNullOrEmpty(definedName.Text) && definedName.Text.StartsWith("temp_"))
                {
                    namesToRemove.Add(definedName.Text);
                }
            }

            // Remove the collected names
            foreach (string name in namesToRemove)
            {
                nameCollection.Remove(name);
            }

            // Save the workbook after filtering
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
