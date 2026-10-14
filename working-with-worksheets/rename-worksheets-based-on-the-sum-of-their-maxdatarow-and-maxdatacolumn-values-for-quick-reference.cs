// Title: Rename Excel worksheets to unique names derived from MaxDataRow and MaxDataColumn using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a workbook with Aspose.Cells, computes MaxDataRow + MaxDataColumn for each worksheet, renames the sheet to "Sheet_<sum>" while guaranteeing uniqueness, and saves the updated file. | Create a .NET snippet that iterates over all worksheets, calculates the sum of the last data row and column indices, applies a unique name based on that sum, and writes the modified workbook back to disk. | Write a program using Aspose.Cells that checks for a missing input file, renames worksheets according to their data extents, adds numeric suffixes when a name collides, and handles CellsException gracefully.
// Common Searches: aspnet rename excel worksheets based on max data row and column sum | c# Aspose.Cells generate unique worksheet names from data dimensions | how to avoid duplicate sheet names when renaming with Aspose.Cells | example code to rename worksheets using MaxDataRow + MaxDataColumn in .NET
// Tags: Aspose.Cells worksheet renaming using MaxDataRow | unique sheet name generation from data extents | C# sum of MaxDataRow and MaxDataColumn | duplicate sheet name avoidance Aspose.Cells | save workbook after renaming Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The program loads an input XLSX file with Aspose.Cells, iterates through each worksheet, calculates the sum of its MaxDataRow and MaxDataColumn indices, builds a unique name like "Sheet_<sum>" (adding a numeric suffix if needed), renames the worksheet, and saves the result to a new file while handling missing files and Aspose.Cells exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Keep track of worksheet names to ensure uniqueness
            HashSet<string> usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // MaxDataRow and MaxDataColumn are zero‑based indices of the last cell that contains data
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                // Calculate the sum of the two values
                int sum = maxRow + maxCol;

                // Build a base name and ensure it is unique within the workbook
                string baseName = $"Sheet_{sum}";
                string newName = baseName;
                int suffix = 1;
                while (usedNames.Contains(newName))
                {
                    newName = $"{baseName}_{suffix}";
                    suffix++;
                }

                // Rename the worksheet
                sheet.Name = newName;
                usedNames.Add(newName);
            }

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
