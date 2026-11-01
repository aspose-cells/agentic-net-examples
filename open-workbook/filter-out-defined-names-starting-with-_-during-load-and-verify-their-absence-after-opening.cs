// Title: Remove underscore‑prefixed defined names from an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# remove defined names that start with underscore | filter out named ranges beginning with '_' when loading a workbook with Aspose.Cells | verify that no underscore‑prefixed named ranges exist after opening an Excel file in C# | how to delete hidden named ranges in an Excel workbook using Aspose.Cells .NET | remove and validate underscore‑prefixed named ranges before saving a workbook
// Tags: Aspose.Cells remove underscore named ranges | C# filter defined names on workbook load | Aspose.Cells validate named range cleanup | Excel hidden named range deletion .NET | named range removal before workbook save

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Loads an Excel file, collects all defined names that begin with an underscore, removes them from the workbook, verifies none remain, and saves the result to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook from the input file
            var workbook = new Workbook(inputPath);

            // Collect defined names that start with "_" for removal
            var namesToRemove = new List<Name>();
            foreach (Name name in workbook.Worksheets.Names)
            {
                if (name.Text.StartsWith("_", StringComparison.Ordinal))
                    namesToRemove.Add(name);
            }

            // Remove the collected names from the workbook
            foreach (Name name in namesToRemove)
            {
                workbook.Worksheets.Names.Remove(name.Text);
            }

            // Verify that no defined name starting with "_" remains
            foreach (Name name in workbook.Worksheets.Names)
            {
                if (name.Text.StartsWith("_", StringComparison.Ordinal))
                    throw new InvalidOperationException($"Defined name '{name.Text}' was not removed.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
