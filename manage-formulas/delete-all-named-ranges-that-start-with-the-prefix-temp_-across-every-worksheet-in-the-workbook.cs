// Title: How to delete all named ranges that start with 'Temp_' from every worksheet using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells remove named ranges with prefix Temp_ from workbook | Delete temporary named ranges across all sheets using Aspose.Cells .NET | Aspose.Cells iterate workbook names collection to filter and delete Temp_* ranges | Programmatically clear named ranges beginning with Temp_ in Excel via Aspose.Cells | Bulk delete of prefixed named ranges in Aspose.Cells C# example
// Tags: named range prefix removal Aspose.Cells | workbook names collection filtering C# | bulk delete of Excel named ranges .NET | temporary named range cleanup Aspose.Cells | Aspose.Cells named range management

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, collects all defined names whose text starts with "Temp_", removes those named ranges from the workbook's Names collection, and saves the modified file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Collect names that start with the prefix "Temp_"
            List<string> namesToDelete = new List<string>();
            foreach (Name name in workbook.Worksheets.Names)
            {
                if (name.Text.StartsWith("Temp_", StringComparison.Ordinal))
                {
                    namesToDelete.Add(name.Text);
                }
            }

            // Delete the collected named ranges
            foreach (string name in namesToDelete)
            {
                workbook.Worksheets.Names.Remove(name);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (CellsException ex)
        {
            // Handles Aspose.Cells specific errors (e.g., password protection)
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handles any other unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
