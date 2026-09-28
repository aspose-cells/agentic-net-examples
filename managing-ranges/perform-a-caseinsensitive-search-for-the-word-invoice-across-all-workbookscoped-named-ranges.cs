// Title: Case‑insensitive search for a keyword across all workbook‑scoped named ranges using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that scans every workbook‑scoped named range and returns the range name, reference address, and the first cell containing the word "invoice" (case‑insensitive). | Adapt the example to collect all cells that contain the search term within each named range instead of stopping after the first match. | Create a reusable method that accepts a workbook file path and a search term, then outputs a list of named ranges with matching cells using Aspose.Cells.
// Common Searches: aspocells c# find text in workbook scoped named ranges case insensitive | how to list Excel named ranges that contain a specific word using Aspose.Cells | search for keyword in all named ranges of an .xlsx file with Aspose.Cells .NET | retrieve cell address of matching text inside a named range using Aspose.Cells C#
// Tags: case-insensitive search in workbook scoped named ranges Aspose.Cells | enumerate named ranges and extract matching cells C# | keyword lookup within Excel named ranges using Aspose.Cells | retrieve matching cell address from named range Aspose.Cells | Aspose.Cells .xlsx named range text search

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The program loads "input.xlsx", iterates through each workbook‑scoped named range, scans its cells for the string "invoice" using a case‑insensitive comparison, and records the named range name, its reference address, and the first matching cell's address.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            string searchTerm = "invoice";
            var matches = new List<string>();

            // Iterate over all workbook‑scoped named ranges
            foreach (Name name in workbook.Worksheets.Names)
            {
                // Get the range that the name refers to
                Aspose.Cells.Range range = name.GetRange();

                // Scan each cell in the range
                foreach (Cell cell in range)
                {
                    if (cell.Type == CellValueType.IsString)
                    {
                        string cellText = cell.StringValue;
                        if (cellText.IndexOf(searchTerm, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            matches.Add($"Name: {name.Text}, Address: {range.RefersTo}, Cell: {cell.Name}");
                            break; // stop scanning this range after the first match
                        }
                    }
                }
            }

            // Output the results
            Console.WriteLine("Workbook‑scoped named ranges containing the word \"invoice\" (case‑insensitive):");
            foreach (string result in matches)
            {
                Console.WriteLine(result);
            }

            // Save the workbook if further processing is required
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
