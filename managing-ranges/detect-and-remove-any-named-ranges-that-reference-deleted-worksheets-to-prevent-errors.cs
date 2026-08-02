// Title: C# – Remove Invalid Named Ranges That Reference Deleted Worksheets Using Aspose.Cells
// Description: Loads a workbook (or creates a new one), optionally deletes a worksheet, scans the NameCollection for defined names whose RefersTo formula points to a missing sheet, removes those broken named ranges, and saves the cleaned file. Prevents #REF! errors after worksheet removal.
// Keywords: Aspose.Cells remove broken named ranges | C# delete named range missing sheet | invalid defined names Aspose.Cells | clean workbook after sheet deletion | Aspose.Cells NameCollection cleanup
// Common Searches: how to delete named ranges that point to deleted worksheets in Aspose.Cells | remove invalid defined names after removing a sheet .NET | Aspose.Cells detect named ranges with missing sheet references | C# clean up named ranges after worksheet removal
// Developer Intent: Find and delete any named ranges that reference worksheets that have been removed, ensuring the workbook remains error‑free.
// Use Cases: Sanitize a workbook after programmatically deleting worksheets to avoid #REF! errors caused by stale named ranges. | Validate and cleanse imported workbooks from external sources before further processing. | Prepare workbooks for export or publishing by guaranteeing all defined names point to existing sheets.
// AI Prompts: Generate C# code with Aspose.Cells that scans a workbook and removes named ranges referencing non‑existent worksheets. | Create a method that returns a list of invalid named ranges and deletes them safely in Aspose.Cells. | Explain how to extend the example to log each removed named range to a text file.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook (or creates a new one), optionally deletes a worksheet, scans the NameCollection for defined names whose RefersTo formula points to a missing sheet, removes those broken named ranges, and saves the cleaned file. Prevents #REF! errors after worksheet removal.
    public class RemoveInvalidNamedRanges
    {
        public static void Run()
        {
            try
            {
                // Load workbook if input file exists; otherwise create a new workbook
                string inputPath = "input.xlsx";
                Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

                // Example: delete a worksheet to simulate missing references
                if (workbook.Worksheets.Count > 0)
                {
                    workbook.Worksheets.RemoveAt(0);
                }

                // Get the collection of defined names (named ranges)
                NameCollection names = workbook.Worksheets.Names;

                // List to hold names that reference non‑existent worksheets
                List<string> namesToRemove = new List<string>();

                // Iterate through all defined names
                foreach (Name name in names)
                {
                    if (string.IsNullOrEmpty(name.RefersTo))
                        continue;

                    // RefersTo format is usually like "=SheetName!$A$1:$B$2"
                    string refersTo = name.RefersTo.TrimStart('=');
                    int exclPos = refersTo.IndexOf('!');
                    if (exclPos <= 0)
                        continue; // Not a standard sheet reference

                    string sheetName = refersTo.Substring(0, exclPos);

                    // Check if the worksheet exists in the workbook
                    bool sheetExists = false;
                    foreach (Worksheet ws in workbook.Worksheets)
                    {
                        if (ws.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            sheetExists = true;
                            break;
                        }
                    }

                    // If the sheet does not exist, mark this name for removal
                    if (!sheetExists)
                    {
                        namesToRemove.Add(name.Text);
                    }
                }

                // Remove the invalid named ranges
                if (namesToRemove.Count > 0)
                {
                    names.Remove(namesToRemove.ToArray());
                }

                // Save the workbook
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            RemoveInvalidNamedRanges.Run();
        }
    }
}
