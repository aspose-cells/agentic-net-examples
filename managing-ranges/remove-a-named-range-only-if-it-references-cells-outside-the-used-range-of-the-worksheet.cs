// Title: Remove Named Ranges Outside the Used Area with Aspose.Cells for .NET (C#)
// Description: Loads a workbook, determines the worksheet's used range via MaxDataRow/MaxDataColumn, scans all defined names, identifies those whose referenced range extends beyond the used rows or columns, removes the offending names from the NameCollection, and saves the cleaned file.
// Keywords: Aspose.Cells remove external named ranges | delete named range outside used area | C# Aspose.Cells named range management | filter out‑of‑bounds named ranges | Aspose.Cells .NET clean workbook
// Common Searches: how to delete named ranges that point outside the used range in Aspose.Cells | remove out‑of‑bounds named ranges C# Aspose.Cells | Aspose.Cells check if a named range is beyond MaxDataRow | prune stale named ranges in Excel using Aspose.Cells
// Developer Intent: Programmatically eliminate any named range that references cells beyond the worksheet's populated area.
// Use Cases: Sanitize legacy Excel files before distribution by stripping obsolete named ranges. | Ensure data‑export routines only encounter valid ranges, preventing runtime errors. | Automate workbook validation in CI pipelines to keep file size and complexity low.
// AI Prompts: Generate C# code with Aspose.Cells that logs each removed named range to the console. | Show how to extend the sample to also delete names that refer to whole rows or columns outside the used area. | Create a unit test that confirms named ranges outside the used range are removed after processing.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook, determines the worksheet's used range via MaxDataRow/MaxDataColumn, scans all defined names, identifies those whose referenced range extends beyond the used rows or columns, removes the offending names from the NameCollection, and saves the cleaned file.
    public class RemoveExternalNamedRanges
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Processing completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Load an existing workbook if the file exists; otherwise create a new one.
            string inputPath = "input.xlsx";
            Workbook workbook;

            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a blank workbook
            }

            // Access the first worksheet (adjust if needed)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Determine the used range of the worksheet
            int usedLastRow = cells.MaxDataRow;          // zero‑based index of the last row with data
            int usedLastColumn = cells.MaxDataColumn;    // zero‑based index of the last column with data

            // Collect names that refer to ranges outside the used area
            NameCollection names = workbook.Worksheets.Names;
            List<string> namesToRemove = new List<string>();

            foreach (Name name in names)
            {
                // Only process names that refer to a range
                if (name.RefersTo == null)
                    continue;

                // Get the actual range the name points to
                Aspose.Cells.Range rng = name.GetRange();

                if (rng == null)
                    continue; // not a range reference

                // Calculate the absolute bounds of the range
                int rangeFirstRow = rng.FirstRow;
                int rangeFirstColumn = rng.FirstColumn;
                int rangeLastRow = rangeFirstRow + rng.RowCount - 1;
                int rangeLastColumn = rangeFirstColumn + rng.ColumnCount - 1;

                // Check if any part of the range lies outside the used range
                bool outside = rangeFirstRow > usedLastRow ||
                               rangeFirstColumn > usedLastColumn ||
                               rangeLastRow > usedLastRow ||
                               rangeLastColumn > usedLastColumn;

                if (outside)
                {
                    namesToRemove.Add(name.Text);
                }
            }

            // Remove the identified names
            if (namesToRemove.Count > 0)
            {
                names.Remove(namesToRemove.ToArray());
            }

            // Save the workbook (or to a new file)
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath);
        }
    }
}
