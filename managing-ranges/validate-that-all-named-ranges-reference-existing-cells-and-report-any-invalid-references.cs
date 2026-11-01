// Title: Validate all named ranges in an Excel workbook and report invalid references with Aspose.Cells for .NET (C#)
// Common Searches: how to find named ranges that refer to non‑existent cells using Aspose.Cells in C# | Aspose.Cells C# list undefined named ranges in a workbook | programmatically verify Excel named ranges existence with Aspose.Cells | detect and output invalid named range references in a .NET Excel file | C# code to check that every named range points to a valid cell range
// Tags: named range validation Aspose.Cells | GetRangeByName error handling C# | Excel named range existence check .NET | detect invalid named ranges Aspose.Cells | list undefined named ranges workbook

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, iterates through all defined names, attempts to retrieve each range via GetRangeByName, treats null, empty ranges or any exception as invalid, collects those names, and then prints either a success message or a list of the invalid named ranges.
class NamedRangeValidator
{
    static void Main(string[] args)
    {
        // Path to the workbook to validate
        string filePath = "input.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(filePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Collect names that reference invalid cells
        List<string> invalidNames = new List<string>();

        // Iterate through all defined names in the workbook
        foreach (Name name in workbook.Worksheets.Names)
        {
            try
            {
                // Attempt to retrieve the range object by name
                Aspose.Cells.Range range = workbook.Worksheets.GetRangeByName(name.Text);

                // If the range is null or empty, treat it as invalid
                if (range == null || range.RowCount == 0 || range.ColumnCount == 0)
                {
                    invalidNames.Add(name.Text);
                }
            }
            catch (Exception)
            {
                // Any exception while resolving the reference means it's invalid
                invalidNames.Add(name.Text);
            }
        }

        // Report the validation results
        if (invalidNames.Count == 0)
        {
            Console.WriteLine("All named ranges reference existing cells.");
        }
        else
        {
            Console.WriteLine("Invalid named ranges detected:");
            foreach (string invalidName in invalidNames)
            {
                Console.WriteLine("- " + invalidName);
            }
        }
    }
}
