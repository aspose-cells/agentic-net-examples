// Title: Validate Excel Named Ranges and Detect Invalid References with Aspose.Cells for .NET
// Description: Loads an Excel file, enumerates all defined names, uses Name.GetRanges(true) to resolve each reference, and verifies that the start/end rows and columns fall within the worksheet's populated area (MaxDataRow/MaxDataColumn). Invalid or out‑of‑bounds ranges are logged, and the workbook is saved unchanged.
// Keywords: Aspose.Cells named range validation | C# check Excel named ranges | detect broken named ranges | invalid named range detection | Excel range bounds verification | Aspose.Cells GetRanges example | .NET Excel workbook validation
// Common Searches: how to validate named ranges in Excel using Aspose.Cells | C# code to find invalid named range references | Aspose.Cells check if named range points to existing cells | detect out of bounds named ranges .NET | verify Excel named ranges programmatically
// Developer Intent: Ensure every named range in a workbook references cells that actually exist and report any that are missing or out of bounds.
// Use Cases: Audit workbooks before distribution to prevent formula errors caused by broken named ranges. | Automate cleanup of corrupted named ranges during batch processing of multiple Excel files. | Validate named ranges after programmatic modifications such as row/column deletions or sheet restructuring.
// AI Prompts: Write C# code with Aspose.Cells that iterates all defined names and flags those whose GetRanges returns null or an empty collection. | Create a method that checks each range's FirstRow, FirstColumn, RowCount, and ColumnCount against the worksheet's MaxDataRow and MaxDataColumn, returning a list of invalid named ranges. | Suggest enhancements to automatically remove or correct invalid named ranges after they are detected.

using System;
using System.IO;
using Aspose.Cells;

namespace NamedRangeValidator
{
    // Loads an Excel file, enumerates all defined names, uses Name.GetRanges(true) to resolve each reference, and verifies that the start/end rows and columns fall within the worksheet's populated area (MaxDataRow/MaxDataColumn). Invalid or out‑of‑bounds ranges are logged, and the workbook is saved unchanged.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input and output workbooks
            string inputPath = "input.xlsx";
            string outputPath = "validated.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            Workbook workbook = null;
            try
            {
                // Load the workbook from the specified file (constructor loads the file)
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Collection of all defined names (named ranges) in the workbook
            NameCollection names = workbook.Worksheets.Names;

            bool anyInvalid = false;

            // Iterate through each defined name
            foreach (Name name in names)
            {
                // Obtain the ranges the name refers to (recalculate to get latest info)
                Aspose.Cells.Range[] ranges = name.GetRanges(true);

                // If GetRanges returns null or an empty array, the reference is likely invalid
                if (ranges == null || ranges.Length == 0)
                {
                    Console.WriteLine($"Invalid reference detected: Name '{name.Text}' has no resolvable ranges. RefersTo = {name.RefersTo}");
                    anyInvalid = true;
                    continue;
                }

                // Validate each range returned
                foreach (Aspose.Cells.Range range in ranges)
                {
                    Worksheet ws = range.Worksheet;

                    // Determine the maximum occupied row/column in the worksheet
                    int maxRow = ws.Cells.MaxDataRow;      // -1 if worksheet is empty
                    int maxCol = ws.Cells.MaxDataColumn;   // -1 if worksheet is empty

                    // Validate start and end indices against worksheet limits
                    bool startRowValid = range.FirstRow >= 0 && (maxRow == -1 || range.FirstRow <= maxRow);
                    bool endRowValid   = range.RowCount > 0 && (maxRow == -1 || (range.FirstRow + range.RowCount - 1) <= maxRow);
                    bool startColValid = range.FirstColumn >= 0 && (maxCol == -1 || range.FirstColumn <= maxCol);
                    bool endColValid   = range.ColumnCount > 0 && (maxCol == -1 || (range.FirstColumn + range.ColumnCount - 1) <= maxCol);

                    if (!startRowValid || !endRowValid || !startColValid || !endColValid)
                    {
                        Console.WriteLine($"Invalid reference detected: Name '{name.Text}' refers to range '{range.RefersTo}' which is out of bounds in worksheet '{ws.Name}'.");
                        anyInvalid = true;
                    }
                }
            }

            if (!anyInvalid)
            {
                Console.WriteLine("All named ranges reference valid cells.");
            }

            // Ensure the output directory exists
            try
            {
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the (unchanged) workbook to the output path
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
