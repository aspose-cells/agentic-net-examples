// Title: Load only the "Summary" and "Data" worksheets from an Excel file and discard other sheets using Aspose.Cells in C#
// AI Prompts: Write C# code that opens an XLSX file with Aspose.Cells, retains only the worksheets named "Summary" and "Data", removes all others, and saves the filtered workbook. | Show how to implement a case‑insensitive whitelist of sheet names when processing a workbook with Aspose.Cells, then export the cleaned file. | Create a reusable method that accepts a file path and a collection of sheet names, returns a new Workbook containing only those sheets using Aspose.Cells.
// Common Searches: Aspose.Cells keep only selected worksheets C# | How to remove unwanted Excel sheets after loading with Aspose.Cells | C# whitelist specific worksheet names when saving workbook Aspose.Cells | Example of discarding all sheets except Summary and Data using Aspose.Cells | Aspose.Cells filter workbook sheets on load without LoadFilter
// Tags: retain specific worksheets Aspose.Cells | remove unwanted sheets C# Aspose.Cells | case‑insensitive sheet whitelist Aspose.Cells | filter workbook sheets after load Aspose.Cells | save filtered Excel file Aspose.Cells | C# Excel sheet selection Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The program checks that the input XLSX file exists, loads it into a Workbook, defines a case‑insensitive set containing "Summary" and "Data", iterates the worksheets collection backward to remove any sheet not in the set, and then saves the filtered workbook to a new file.
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
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook without any filter
            Workbook workbook = new Workbook(inputPath);

            // Define the sheets we want to keep (case‑insensitive)
            HashSet<string> allowedSheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "Summary",
                "Data"
            };

            // Remove sheets that are not in the allowed list
            // Iterate backwards because removing sheets changes the collection index
            for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
            {
                Worksheet sheet = workbook.Worksheets[i];
                if (!allowedSheets.Contains(sheet.Name))
                {
                    workbook.Worksheets.RemoveAt(i);
                }
            }

            // Save the filtered workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
