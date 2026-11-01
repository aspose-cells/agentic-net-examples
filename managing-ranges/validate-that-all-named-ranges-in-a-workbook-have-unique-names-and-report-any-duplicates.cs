// Title: Check for duplicate named ranges in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells find duplicate named ranges in an existing workbook | how to list repeated defined names using Aspose.Cells .NET | detect non‑unique named ranges in Excel file with Aspose.Cells API | Aspose.Cells C# code to report named range name collisions | validate named range uniqueness case insensitive Aspose.Cells
// Tags: Aspose.Cells named range duplication detection | C# named range uniqueness check | Excel workbook defined name collision detection | case-insensitive name counting with Aspose.Cells | reporting duplicate defined names Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads (or creates) an Excel workbook, iterates through all defined names via workbook.Worksheets.Names, counts each name case‑insensitively using a dictionary, gathers any names that appear more than once, and writes a console report listing each duplicate name with its occurrence count.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            Workbook workbook = null;

            try
            {
                if (File.Exists(inputPath))
                {
                    // Load existing workbook (add LoadOptions.Password if needed)
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    // Create a new workbook if the file does not exist
                    workbook = new Workbook();
                    workbook.Worksheets[0].Name = "Sheet1";
                    workbook.Save(inputPath);
                    Console.WriteLine($"Input file not found. Created a new workbook at '{inputPath}'.");
                }
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Aspose.Cells error while loading workbook: {ex.Message}");
                return;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
                return;
            }

            // Track named range occurrences (case‑insensitive)
            var nameCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var duplicateNames = new List<string>();

            // Iterate through all named ranges in the workbook
            foreach (Name namedRange in workbook.Worksheets.Names)
            {
                string name = namedRange.Text;

                if (nameCounts.ContainsKey(name))
                {
                    nameCounts[name]++;
                    if (nameCounts[name] == 2)
                    {
                        duplicateNames.Add(name);
                    }
                }
                else
                {
                    nameCounts[name] = 1;
                }
            }

            // Report duplicate named ranges
            if (duplicateNames.Count > 0)
            {
                Console.WriteLine("Duplicate named ranges found:");
                foreach (string dup in duplicateNames)
                {
                    Console.WriteLine($"- {dup} (occurs {nameCounts[dup]} times)");
                }
            }
            else
            {
                Console.WriteLine("All named ranges have unique names.");
            }

            // Optional: save the workbook if modifications were made
            // workbook.Save("output.xlsx");
        }
    }
}
