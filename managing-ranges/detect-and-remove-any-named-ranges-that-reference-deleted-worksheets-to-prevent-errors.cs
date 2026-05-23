using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RemoveInvalidNamedRanges
    {
        public static void Run()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Ensure the input file exists; create a simple workbook if it does not.
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    workbook.Worksheets[0].Name = "Sheet1";
                    workbook.Save(inputPath);
                }

                // Example: delete a worksheet to create an invalid reference scenario
                if (workbook.Worksheets.Count > 1)
                {
                    workbook.Worksheets.RemoveAt(1); // delete second worksheet (index 1)
                }

                // Build a hash set of current worksheet names for fast lookup
                var existingSheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    existingSheets.Add(ws.Name);
                }

                // Collect names that reference non‑existent worksheets
                var namesToRemove = new List<string>();
                NameCollection names = workbook.Worksheets.Names;

                foreach (Name name in names)
                {
                    // Skip if the name does not refer to a range/formula
                    if (string.IsNullOrEmpty(name.RefersTo))
                        continue;

                    // RefersTo string starts with '='; remove it for parsing
                    string refers = name.RefersTo.TrimStart('=');

                    // Split multiple references separated by commas
                    string[] parts = refers.Split(',');

                    foreach (string part in parts)
                    {
                        // Find the sheet name before the first '!' character
                        int exclPos = part.IndexOf('!');
                        if (exclPos > 0)
                        {
                            string sheetName = part.Substring(0, exclPos).Trim('\''); // remove possible surrounding quotes
                            // If the sheet name is not in the current workbook, mark the name for removal
                            if (!existingSheets.Contains(sheetName))
                            {
                                namesToRemove.Add(name.Text);
                                break; // no need to check other parts for this name
                            }
                        }
                    }
                }

                // Remove the invalid named ranges
                if (namesToRemove.Count > 0)
                {
                    names.Remove(namesToRemove.ToArray());
                    Console.WriteLine("Removed invalid named ranges:");
                    foreach (string n in namesToRemove)
                    {
                        Console.WriteLine($"- {n}");
                    }
                }
                else
                {
                    Console.WriteLine("No invalid named ranges found.");
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point required for console application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}