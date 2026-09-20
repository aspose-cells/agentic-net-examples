// Title: Validate named ranges in an Excel workbook with Aspose.Cells for .NET and list any out‑of‑bounds references
// AI Prompts: Generate a C# method that loads a workbook, iterates through workbook.Worksheets.Names, calls Name.GetRange(), compares the range coordinates with the worksheet's MaxDataRow/MaxDataColumn, and returns the identifiers of named ranges that fall outside the sheet. | Write code that examines each defined name for a missing worksheet or a range that exceeds the sheet's data bounds and prints a formatted report of all invalid named range names.
// Common Searches: how to programmatically find invalid named ranges in an Excel file using Aspose.Cells C# | C# Aspose.Cells example to list named ranges that point to non‑existent cells | detect out‑of‑range named range references in a .xlsx workbook with Aspose.Cells | validate defined names against worksheet data limits using Aspose.Cells for .NET | Aspose.Cells GetRange throws exception for missing worksheet named range
// Tags: Aspose.Cells validate named ranges | C# check named range bounds | Aspose.Cells GetRange out-of-bounds detection | Excel workbook named range verification .NET | list invalid defined names Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads an Excel file, iterates through all defined names, obtains each range via Name.GetRange(), verifies that the referenced worksheet exists and that the range's start and end rows/columns are within the worksheet's maximum data rows and columns, collects any names with invalid references, and prints a concise report of those invalid named ranges.
public class NamedRangeValidator
{
    /// <param name="filePath">Path to the Excel file to validate.</param>
    public static void ValidateNamedRanges(string filePath)
    {
        // Ensure the file exists before attempting to load it.
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        Workbook workbook;
        try
        {
            workbook = new Workbook(filePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        List<string> invalidNames = new List<string>();

        // Iterate through all defined names (named ranges) in the workbook.
        foreach (Name definedName in workbook.Worksheets.Names)
        {
            try
            {
                // Attempt to get the range that the name refers to.
                AsposeRange range = definedName.GetRange();

                // Validate that the range's worksheet is available.
                Worksheet ws = range.Worksheet;
                if (ws == null)
                {
                    invalidNames.Add(definedName.Text); // Use Text as fallback for the name.
                    continue;
                }

                // Determine worksheet bounds.
                int maxRow = ws.Cells.MaxDataRow;
                int maxCol = ws.Cells.MaxDataColumn;

                // If worksheet is empty, treat bounds as -1; any range is invalid.
                if (maxRow < 0 || maxCol < 0)
                {
                    invalidNames.Add(definedName.Text);
                    continue;
                }

                // Calculate range boundaries.
                int startRow = range.FirstRow;
                int endRow = range.FirstRow + range.RowCount - 1;
                int startCol = range.FirstColumn;
                int endCol = range.FirstColumn + range.ColumnCount - 1;

                // Validate the range coordinates against worksheet bounds.
                if (startRow < 0 || endRow > maxRow ||
                    startCol < 0 || endCol > maxCol)
                {
                    invalidNames.Add(definedName.Text);
                }
            }
            catch (Exception)
            {
                // If an exception occurs, the reference is invalid.
                invalidNames.Add(definedName.Text);
            }
        }

        // Report results.
        if (invalidNames.Count == 0)
        {
            Console.WriteLine("All named ranges reference valid cells.");
        }
        else
        {
            Console.WriteLine("Invalid named ranges found:");
            foreach (string invalidName in invalidNames)
            {
                Console.WriteLine($"- {invalidName}");
            }
        }
    }

    // Entry point for the console application.
    public static void Main(string[] args)
    {
        // Example usage: provide the Excel file path as the first argument.
        if (args.Length == 0)
        {
            Console.WriteLine("Please provide the path to the Excel file as an argument.");
            return;
        }

        string filePath = args[0];
        ValidateNamedRanges(filePath);
    }
}
