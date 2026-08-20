// Title: C# – Remove Named Ranges That Reference Deleted Worksheets with Aspose.Cells
// Description: This C# example loads an Excel workbook, optionally deletes a worksheet, scans all defined names, identifies those whose RefersTo points to a missing sheet, removes the invalid named ranges, and saves the cleaned file, preventing runtime errors.
// Keywords: Aspose.Cells | C# | named range cleanup | invalid named ranges | deleted worksheet | Excel workbook | remove defined names | RefersTo | worksheet removal
// Common Searches: How to remove named ranges that reference a deleted sheet in Aspose.Cells C# | Aspose.Cells remove invalid defined names after worksheet deletion | C# code to clean up stale named ranges in an Excel workbook | Detect and delete named ranges pointing to non‑existent worksheets Aspose.Cells | Remove named ranges referencing external workbooks with Aspose.Cells
// Developer Intent: Find and delete any named ranges that refer to worksheets that have been removed from the workbook.
// Use Cases: After programmatically deleting a worksheet, clean up the workbook by removing named ranges that still reference the deleted sheet to avoid runtime errors. | Validate a workbook before saving by scanning all defined names and discarding those whose RefersTo points to a missing sheet. | Prepare an exported Excel file for third‑party consumption by ensuring no stale named ranges remain after sheet removal.
// AI Prompts: Write C# code using Aspose.Cells to find and remove named ranges that reference worksheets that no longer exist. | Provide a method that returns a list of invalid named ranges after a worksheet deletion in an Aspose.Cells workbook. | Explain how to safely delete a worksheet and clean up associated named ranges without affecting external references.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This C# example loads an Excel workbook, optionally deletes a worksheet, scans all defined names, identifies those whose RefersTo points to a missing sheet, removes the invalid named ranges, and saves the cleaned file, preventing runtime errors.
    public class RemoveInvalidNamedRanges
    {
        public static void Run()
        {
            try
            {
                string inputPath = "input.xlsx";

                // Ensure the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Delete the third worksheet if it exists
                if (workbook.Worksheets.Count > 2)
                {
                    workbook.Worksheets.RemoveAt(2);
                }

                // Get all defined names (named ranges)
                NameCollection names = workbook.Worksheets.Names;
                List<string> namesToRemove = new List<string>();

                // Identify names that reference non‑existent worksheets
                foreach (Name name in names)
                {
                    if (string.IsNullOrEmpty(name.RefersTo))
                        continue;

                    // RefersTo format: "=Sheet1!$A$1:$B$5"
                    string refersTo = name.RefersTo.TrimStart('=');
                    string[] parts = refersTo.Split('!');

                    if (parts.Length < 2)
                        continue;

                    string sheetName = parts[0];

                    // Skip external workbook references
                    if (sheetName.StartsWith("["))
                        continue;

                    Worksheet ws = workbook.Worksheets[sheetName];
                    if (ws == null)
                    {
                        namesToRemove.Add(name.Text);
                    }
                }

                // Remove invalid named ranges
                if (namesToRemove.Count > 0)
                {
                    names.Remove(namesToRemove.ToArray());
                    Console.WriteLine("Removed the following invalid named ranges:");
                    foreach (string n in namesToRemove)
                    {
                        Console.WriteLine($"- {n}");
                    }
                }
                else
                {
                    Console.WriteLine("No invalid named ranges were found.");
                }

                // Save the modified workbook
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
