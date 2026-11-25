// Title: Detect merged cells in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells check if worksheet has merged cells | Aspose.Cells .NET find merged ranges in Excel file | How to list merged cell areas using Aspose.Cells C# | Detect merged cells without scanning the entire sheet Aspose.Cells | Get merged range reference Aspose.Cells C# example
// Tags: Aspose.Cells detect merged ranges | C# check merged cells Excel | Worksheet.Cells IsMerged property | GetMergedRange method usage | Iterate used range Aspose.Cells | Excel merged area detection .NET

using System;
using System.IO;
using Aspose.Cells;

namespace CheckMergedCellsDemo
{
    // The sample loads a workbook (or creates one with a merged B2:D4 range), selects the first worksheet, determines the used row and column limits, and iterates through each cell within that range. For each cell it evaluates the IsMerged flag; when true it obtains the full merged range via GetMergedRange and outputs the range reference. Finally, it reports whether any merged cells were found.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Load or create a workbook
                // -----------------------------------------------------------------
                Workbook workbook;
                string inputPath = args.Length > 0 ? args[0] : null;

                if (!string.IsNullOrEmpty(inputPath) && File.Exists(inputPath))
                {
                    // Load the workbook from the provided file path
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook from \"{inputPath}\".");
                }
                else
                {
                    if (!string.IsNullOrEmpty(inputPath))
                    {
                        Console.WriteLine($"File \"{inputPath}\" not found. Creating an in‑memory workbook instead.");
                    }

                    // Create a sample workbook with a merged range for demonstration
                    workbook = new Workbook();
                    Worksheet ws = workbook.Worksheets[0];
                    ws.Name = "SampleSheet";

                    // Merge cells B2:D4
                    ws.Cells.CreateRange("B2", "D4").Merge();
                    // Put a value into the merged area
                    ws.Cells["B2"].PutValue("Merged Area");

                    Console.WriteLine("Created an in‑memory workbook with a merged range (B2:D4).");
                }

                // -----------------------------------------------------------------
                // 2. Get the worksheet to inspect (first worksheet)
                // -----------------------------------------------------------------
                Worksheet worksheet = workbook.Worksheets[0];
                Console.WriteLine($"Inspecting worksheet \"{worksheet.Name}\"...");

                // -----------------------------------------------------------------
                // 3. Determine the used range limits to avoid iterating over the entire grid
                // -----------------------------------------------------------------
                int maxRow = worksheet.Cells.MaxDataRow;      // -1 if worksheet is empty
                int maxCol = worksheet.Cells.MaxDataColumn;   // -1 if worksheet is empty

                if (maxRow < 0 || maxCol < 0)
                {
                    Console.WriteLine("Worksheet contains no data cells – therefore no merged cells.");
                    return;
                }

                bool hasMergedCells = false;

                // -----------------------------------------------------------------
                // 4. Scan the cells collection for merged cells
                // -----------------------------------------------------------------
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = worksheet.Cells[row, col];

                        // The IsMerged property is true for any cell that belongs to a merged range.
                        if (cell.IsMerged)
                        {
                            // Retrieve the full merged range that this cell belongs to.
                            Aspose.Cells.Range mergedRange = cell.GetMergedRange();
                            if (mergedRange != null &&
                                mergedRange.FirstRow == row && mergedRange.FirstColumn == col)
                            {
                                hasMergedCells = true;
                                Console.WriteLine($"Merged range found: {mergedRange.RefersTo}");
                            }
                        }
                    }
                }

                // -----------------------------------------------------------------
                // 5. Final result output
                // -----------------------------------------------------------------
                if (!hasMergedCells)
                {
                    Console.WriteLine("No merged cells were found in the worksheet.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
