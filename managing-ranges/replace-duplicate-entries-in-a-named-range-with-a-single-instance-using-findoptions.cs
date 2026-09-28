// Title: Remove duplicate values from a named range in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that opens an existing .xlsx file, accesses a specified named range, and clears any duplicate cell values while keeping the first occurrence. | Create a reusable C# method that accepts input file path, output file path, and named range name, then uses Aspose.Cells to iterate the range and eliminate duplicates with a HashSet or FindOptions.
// Common Searches: c# aspose.cells how to delete duplicate entries in a specific named range | remove duplicate cells from an Excel named range using Aspose.Cells library | aspose.cells find and clear duplicate values in a defined range | deduplicate values in Excel named range programmatically with C#
// Tags: Aspose.Cells deduplicate named range C# | clear duplicate cells in Excel range Aspose.Cells | hashset based duplicate detection Aspose.Cells | C# remove repeated values from Excel named range | Aspose.Cells FindOptions for duplicate elimination

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, locates a named range, iterates through its cells, clears any repeated values while preserving the first occurrence, and saves the result to a new file.
class ReplaceDuplicatesInNamedRange
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string rangeName = "MyRange";

            // Verify the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range definition
            Name namedRange = workbook.Worksheets.Names[rangeName];
            if (namedRange == null)
            {
                Console.WriteLine($"Named range \"{rangeName}\" not found.");
                return;
            }

            // Remove leading '=' if present
            string refersTo = namedRange.RefersTo.TrimStart('=');

            // Determine sheet name and address
            string sheetName;
            string rangeAddress;
            if (refersTo.Contains('!'))
            {
                var parts = refersTo.Split('!');
                sheetName = parts[0].Trim('\''); // remove possible quotes
                rangeAddress = parts[1];
            }
            else
            {
                // Fallback to the first worksheet if sheet name is not specified
                sheetName = workbook.Worksheets[0].Name;
                rangeAddress = refersTo;
            }

            // Get the worksheet referenced by the named range
            Worksheet sheet = workbook.Worksheets[sheetName];
            if (sheet == null)
            {
                Console.WriteLine($"Worksheet \"{sheetName}\" not found.");
                return;
            }

            // Create a range object for iteration
            Aspose.Cells.Range range = sheet.Cells.CreateRange(rangeAddress);

            // Track seen values (case‑insensitive)
            HashSet<string> seenValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Remove duplicates within the named range
            foreach (Cell cell in range)
            {
                string cellValue = cell.StringValue ?? string.Empty;

                if (seenValues.Contains(cellValue))
                {
                    cell.PutValue(string.Empty); // Clear duplicate
                }
                else
                {
                    seenValues.Add(cellValue);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Processing complete. Output saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
