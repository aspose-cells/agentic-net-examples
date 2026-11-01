// Title: Populate an Excel template with Aspose.Cells in C# while skipping every other row (alternating rows)
// AI Prompts: Generate C# code that loads a .xlsx template with Aspose.Cells, iterates over a list, and writes each item to column A on every second row starting after the header. | Show how to compute the destination row index for alternating rows when populating a worksheet using Aspose.Cells in C#. | Demonstrate creating the output folder if missing and saving the modified workbook to a new file with Aspose.Cells.
// Common Searches: aspocells c# write data to every other row in existing excel template | how to skip rows when populating worksheet using Aspose.Cells C# | c# aspocells populate list into alternating rows starting after header | aspocells calculate target row index for spaced rows c# example | save modified workbook to new file using Aspose.Cells C#
// Tags: alternating row insertion aspocells c# | skip rows aspocells c# | target row index calculation aspocells c# | template workbook access aspocells c# | save populated workbook aspocells c#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSkipRowsExample
{
    // The example loads an existing .xlsx template (or creates a new workbook), iterates over a list of strings, and writes each entry to column A of the first worksheet with a two‑row gap between entries, starting after the header row. It ensures the output directory exists and saves the populated workbook to a specified path.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing Excel template
            string templatePath = @"C:\Templates\MyTemplate.xlsx";

            // Path where the populated workbook will be saved
            string outputPath = @"C:\Output\PopulatedWorkbook.xlsx";

            try
            {
                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                Workbook workbook;

                // Load the template if it exists; otherwise create a new workbook
                if (File.Exists(templatePath))
                {
                    workbook = new Workbook(templatePath);
                }
                else
                {
                    Console.WriteLine($"Template not found at '{templatePath}'. Creating a new workbook.");
                    workbook = new Workbook(); // creates a workbook with a default worksheet
                }

                // Get the first worksheet (or specify by name/index)
                Worksheet sheet = workbook.Worksheets[0];

                // Starting row index (0‑based). Adjust if the template has headers.
                int startRow = 1; // e.g., start after header row

                // Sample data to populate
                List<string> entries = new List<string>
                {
                    "Entry A",
                    "Entry B",
                    "Entry C",
                    "Entry D",
                    "Entry E"
                };

                // Populate the worksheet, skipping every other row
                for (int i = 0; i < entries.Count; i++)
                {
                    // Calculate the target row: startRow + (i * 2) to skip one row each time
                    int targetRow = startRow + (i * 2);

                    // Write the entry into column A (index 0)
                    Cell cell = sheet.Cells[targetRow, 0];
                    cell.PutValue(entries[i]);

                    // Optionally, you can populate additional columns here
                    // sheet.Cells[targetRow, 1].PutValue("Additional data");
                }

                // Save the populated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
