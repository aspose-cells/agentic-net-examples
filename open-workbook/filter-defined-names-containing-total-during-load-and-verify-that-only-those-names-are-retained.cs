// Title: How to filter and keep only defined names containing "Total" in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# filter defined names by keyword during workbook load | remove Excel named ranges that do not contain "Total" using Aspose.Cells | retain only named ranges with specific text in Aspose.Cells .NET example | case-insensitive removal of unwanted defined names in an Aspose.Cells workbook | verify that all remaining defined names contain a keyword after filtering with Aspose.Cells
// Tags: Aspose.Cells defined name filtering | C# remove unwanted Excel named ranges | case-insensitive named range selection .NET | validate retained defined names after workbook load | retain named ranges containing specific text

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, iterates the workbook's defined name collection in reverse, removes any name whose text does not contain the word "Total" (case‑insensitive), checks that all remaining names include "Total", and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file \"{inputPath}\" was not found.");

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Access the collection of defined names via Worksheets
            NameCollection definedNames = workbook.Worksheets.Names;

            // Remove names that do NOT contain "Total" (case‑insensitive) by iterating backwards
            for (int i = definedNames.Count - 1; i >= 0; i--)
            {
                Name nameObj = definedNames[i];
                if (!nameObj.Text.Contains("Total", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        definedNames.RemoveAt(i);
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine($"Failed to remove defined name at index {i}: {ex.Message}");
                    }
                }
            }

            // Verify that only names containing "Total" remain
            foreach (Name definedName in definedNames)
            {
                if (!definedName.Text.Contains("Total", StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        $"A defined name does not contain \"Total\". Text: \"{definedName.Text}\"");
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Log or display the error as needed
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
