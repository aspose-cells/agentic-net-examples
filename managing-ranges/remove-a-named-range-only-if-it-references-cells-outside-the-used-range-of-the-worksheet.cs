// Title: C# – Remove named ranges that reference cells outside the worksheet’s used range with Aspose.Cells
// Common Searches: how to delete defined names that point outside the used range using Aspose.Cells in C# | Aspose.Cells C# remove named ranges beyond MaxDataRow and MaxDataColumn | filter out-of-bounds named ranges from an Excel workbook with Aspose.Cells | C# code to validate and delete invalid named ranges in a .xlsx file using Aspose.Cells | remove external named ranges that reference non‑existent cells with Aspose.Cells .NET
// Tags: Aspose.Cells delete out-of-bounds defined names | C# check named range limits with MaxDataRow | Excel .xlsx remove invalid named ranges Aspose | named range GetRange error handling Aspose.Cells | worksheet used range detection Aspose.Cells API

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, iterates through all defined names, resolves each to a Range, compares the range boundaries with the worksheet's MaxDataRow and MaxDataColumn, collects indices of names that extend beyond the used area, removes those names, and saves the updated workbook.
class RemoveExternalNamedRanges
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Ensure the input file exists before attempting to load
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Remove named ranges that refer to cells outside the used range
        RemoveInvalidNames(workbook.Worksheets.Names, workbook);

        // Save the modified workbook
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }

    /// <param name="names">Collection of all defined names.</param>
    /// <param name="workbook">Parent workbook.</param>
    private static void RemoveInvalidNames(NameCollection names, Workbook workbook)
    {
        var indicesToRemove = new List<int>();

        for (int i = 0; i < names.Count; i++)
        {
            Name name = names[i];
            Aspose.Cells.Range range;

            // Attempt to resolve the name to a range; skip if it cannot be resolved (e.g., constants)
            try
            {
                range = name.GetRange();
            }
            catch
            {
                continue;
            }

            // Determine the worksheet that the range belongs to; use first sheet as fallback
            Worksheet refSheet = range.Worksheet ?? workbook.Worksheets[0];
            int maxRow = refSheet.Cells.MaxDataRow;      // -1 if the sheet is empty
            int maxCol = refSheet.Cells.MaxDataColumn;   // -1 if the sheet is empty

            bool isOutside = maxRow < 0 || maxCol < 0;

            if (!isOutside)
            {
                int rangeFirstRow = range.FirstRow;
                int rangeFirstCol = range.FirstColumn;
                int rangeLastRow = rangeFirstRow + range.RowCount - 1;
                int rangeLastCol = rangeFirstCol + range.ColumnCount - 1;

                // Check if any part of the range lies outside the used range
                if (rangeFirstRow > maxRow ||
                    rangeFirstCol > maxCol ||
                    rangeLastRow > maxRow ||
                    rangeLastCol > maxCol)
                {
                    isOutside = true;
                }
            }

            if (isOutside)
            {
                indicesToRemove.Add(i);
            }
        }

        // Remove names in reverse order to keep indices valid
        for (int i = indicesToRemove.Count - 1; i >= 0; i--)
        {
            try
            {
                names.RemoveAt(indicesToRemove[i]);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to remove name at index {indicesToRemove[i]}: {ex.Message}");
            }
        }
    }
}
