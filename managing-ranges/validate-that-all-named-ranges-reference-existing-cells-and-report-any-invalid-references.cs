// Title: C# – Validate and Report Invalid Named Ranges with Aspose.Cells
// Description: Loads an Excel file, scans every defined name, uses Name.GetRanges(true) to retrieve its cells, checks that each range and its first cell exist, gathers names with broken references, creates a worksheet called InvalidNamesReport listing them, and saves the updated workbook.
// Keywords: Aspose.Cells named range validation | C# detect broken named ranges | Excel GetRanges invalid reference | .NET named range integrity check | report invalid named ranges Aspose | Excel workbook diagnostics C# | Aspose.Cells GetRanges exception handling
// Common Searches: how to find invalid named ranges using Aspose.Cells | C# code to list broken named ranges in Excel | Aspose.Cells GetRanges returns empty for missing cells | generate report of undefined named ranges .NET | validate named ranges before saving workbook
// Developer Intent: Identify every defined name that points to a non‑existent cell or worksheet and produce a clear list for correction.
// Use Cases: Pre‑publish validation to prevent runtime errors caused by deleted rows or columns. | Automatic creation of a diagnostic sheet that end users can edit to fix broken names. | Integration into CI pipelines that verify Excel data models for reference integrity.
// AI Prompts: Write a C# function with Aspose.Cells that returns a collection of defined names whose referenced ranges are out of bounds. | Generate code that adds a worksheet named "InvalidNamesReport" and populates it with all invalid named range identifiers. | Create a console application that validates named ranges and logs the results to both the console and a new report sheet.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Loads an Excel file, scans every defined name, uses Name.GetRanges(true) to retrieve its cells, checks that each range and its first cell exist, gathers names with broken references, creates a worksheet called InvalidNamesReport listing them, and saves the updated workbook.
class ValidateNamedRanges
{
    static void Main(string[] args)
    {
        // Determine input and output paths
        string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
        string outputPath = args.Length > 1 ? args[1] : "output.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook (lifecycle rule: load)
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // List to hold names that reference non‑existent cells
        List<string> invalidNames = new List<string>();

        // Iterate through every defined name in the workbook
        foreach (Name name in workbook.Worksheets.Names)
        {
            try
            {
                // Retrieve all ranges the name refers to, forcing recalculation
                AsposeRange[] ranges = name.GetRanges(true);

                // If no ranges are returned, the reference is invalid
                if (ranges == null || ranges.Length == 0)
                {
                    invalidNames.Add(name.Text);
                    continue;
                }

                // Validate each returned range
                foreach (AsposeRange rng in ranges)
                {
                    // If the range or its worksheet is null, mark as invalid
                    if (rng == null || rng.Worksheet == null)
                    {
                        invalidNames.Add(name.Text);
                        break;
                    }

                    // Attempt to access the first cell; an exception means the address is out of bounds
                    try
                    {
                        Cell firstCell = rng[0, 0];
                    }
                    catch
                    {
                        invalidNames.Add(name.Text);
                        break;
                    }
                }
            }
            catch
            {
                // Any exception while obtaining ranges indicates a bad reference
                invalidNames.Add(name.Text);
            }
        }

        // Output validation results to the console
        if (invalidNames.Count == 0)
        {
            Console.WriteLine("All named ranges are valid.");
        }
        else
        {
            Console.WriteLine("Invalid named ranges found:");
            foreach (string n in invalidNames)
            {
                Console.WriteLine("- " + n);
            }
        }

        // Optional: create a worksheet that lists the invalid names
        Worksheet reportSheet = workbook.Worksheets[workbook.Worksheets.Add()];
        reportSheet.Name = "InvalidNamesReport";
        int row = 0;
        reportSheet.Cells[row, 0].PutValue("Invalid Named Ranges");
        row++;
        foreach (string n in invalidNames)
        {
            reportSheet.Cells[row, 0].PutValue(n);
            row++;
        }

        // Save the workbook (lifecycle rule: save)
        try
        {
            // Ensure the output directory exists
            string outDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
            {
                Directory.CreateDirectory(outDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
