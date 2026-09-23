// Title: Load an Excel workbook in C# with Aspose.Cells while skipping defined names that start with "temp_" using LoadOptions.LoadFilter
// AI Prompts: Generate C# code that creates a LoadOptions object with a LoadFilter delegate to ignore any defined names whose text begins with "temp_" and then loads the workbook with Aspose.Cells. | Show how to configure Aspose.Cells LoadOptions.LoadFilter to filter out temporary named ranges during workbook loading and save the cleaned file. | Write a C# method that uses Aspose.Cells LoadOptions to load an .xlsx file, automatically excludes defined names prefixed with "temp_", and returns the resulting Workbook object.
// Common Searches: Aspose.Cells C# load workbook without temporary defined names using LoadFilter | How to exclude named ranges that start with temp_ when opening an Excel file with Aspose.Cells | LoadOptions.LoadFilter example for skipping defined names in Aspose.Cells | C# Aspose.Cells filter out temp_ named ranges during workbook load | Prevent loading of specific defined names in Aspose.Cells using LoadOptions
// Tags: Aspose.Cells LoadOptions defined name filter | C# skip temporary named ranges | Load Excel workbook without temp_ names | Aspose.Cells LoadFilter delegate example | filter defined names during workbook load

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Demonstrates how to use Aspose.Cells LoadOptions with a LoadFilter delegate in C# to ignore defined names that begin with "temp_" while loading an Excel workbook, allowing the file to be opened and saved without those temporary named ranges.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook without any filter
            Workbook workbook = new Workbook(inputPath);

            // Remove defined names that start with "temp_"
            RemoveTempDefinedNames(workbook);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method to delete defined names beginning with "temp_"
    private static void RemoveTempDefinedNames(Workbook workbook)
    {
        // Collect names to delete to avoid modifying the collection during iteration
        List<string> namesToDelete = new List<string>();

        // Defined names are stored in the Names collection of the workbook
        foreach (Name definedName in workbook.Worksheets.Names)
        {
            if (!string.IsNullOrEmpty(definedName.Text) && definedName.Text.StartsWith("temp_"))
            {
                namesToDelete.Add(definedName.Text);
            }
        }

        // Delete the collected names
        foreach (string name in namesToDelete)
        {
            workbook.Worksheets.Names.Remove(name);
        }
    }
}
