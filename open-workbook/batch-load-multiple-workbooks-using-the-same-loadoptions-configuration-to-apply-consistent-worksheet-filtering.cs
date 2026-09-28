// Title: Load multiple Excel workbooks with a shared LoadOptions configuration and keep only worksheets that start with a specific prefix using Aspose.Cells for .NET
// AI Prompts: Write C# code that iterates over a list of .xlsx file paths, loads each workbook with a common LoadOptions object, and removes every worksheet whose name does not begin with a given prefix. | Generate a method that receives a collection of workbook paths and a worksheet prefix, opens each file with a shared LoadOptions instance, filters out sheets that don't match the prefix, and returns the count of retained sheets per workbook.
// Common Searches: asp.net load several Excel files with the same LoadOptions and filter sheets by name prefix | c# batch open workbooks using Aspose.Cells and keep only sheets starting with 'Data' | how to apply identical LoadOptions to multiple workbooks in Aspose.Cells | remove worksheets that don't match a naming pattern when loading Excel files in C#
// Tags: batch workbook loading with shared LoadOptions Aspose.Cells | prefix‑based worksheet filtering C# | remove non‑matching sheets during Aspose.Cells import | iterate over multiple .xlsx files using Aspose.Cells | consistent sheet filtering across workbooks C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example iterates through a list of Excel file paths, loads each workbook using a common LoadOptions for Xlsx format, deletes any worksheet whose name does not start with the defined prefix, and outputs the number of retained worksheets while handling missing files and load errors.
class Program
{
    static void Main()
    {
        // List of workbook file paths to be loaded
        List<string> workbookFiles = new List<string>
        {
            "Book1.xlsx",
            "Book2.xlsx",
            "Book3.xlsx"
        };

        // Prefix used to filter worksheets after loading
        const string worksheetPrefix = "Data";

        // Load each workbook safely
        foreach (string filePath in workbookFiles)
        {
            try
            {
                // Ensure the file exists before attempting to load
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                // Load the workbook (full load)
                Workbook workbook = new Workbook(filePath, new LoadOptions(LoadFormat.Xlsx));

                // Remove worksheets that do not start with the specified prefix
                for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
                {
                    Worksheet sheet = workbook.Worksheets[i];
                    if (!sheet.Name.StartsWith(worksheetPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        workbook.Worksheets.RemoveAt(i);
                    }
                }

                // Example usage: display the number of worksheets actually retained
                Console.WriteLine($"Loaded '{filePath}' with {workbook.Worksheets.Count} worksheet(s) matching prefix \"{worksheetPrefix}\".");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors and continue processing other files
                Console.WriteLine($"Error loading '{filePath}': {ex.Message}");
            }
        }
    }
}
