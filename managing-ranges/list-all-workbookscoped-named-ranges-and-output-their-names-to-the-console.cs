// Title: How to enumerate workbook‑scoped named ranges in an Excel file and print them using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells list workbook level defined names from an .xlsx file | retrieve all workbook scoped named ranges using Aspose.Cells .NET | print Excel named ranges to console with Aspose.Cells in C# | enumerate workbook scoped names collection Aspose.Cells example | Aspose.Cells get defined names without worksheet context C#
// Tags: Aspose.Cells workbook.Worksheets.Names usage | C# output Excel defined names to console | load .xlsx workbook with Aspose.Cells .NET | handle missing Excel file Aspose.Cells example | retrieve workbook level names via Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the presence of an input.xlsx file, loads it into an Aspose.Cells Workbook, iterates over the workbook‑scoped defined names through the Worksheets.Names collection, and writes each name's Text property to the console while handling possible loading errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through workbook‑scoped defined names
            // Workbook-level names are accessed via the Worksheets collection's Names property
            foreach (Name name in workbook.Worksheets.Names)
            {
                // Output the name text (e.g., the defined name string)
                Console.WriteLine(name.Text);
            }

            // Optionally, iterate through each worksheet's local names if needed
            /*
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Name localName in sheet.Names)
                {
                    Console.WriteLine($"{sheet.Name}: {localName.Text}");
                }
            }
            */
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions (e.g., loading errors)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
