// Title: Flag Excel worksheets that have a positive MaxDataRow but contain only formatting (no cell values) using Aspose.Cells for .NET
// Common Searches: how to detect worksheets that contain only formatting with Aspose.Cells in C# | Aspose.Cells flag sheets where MaxDataRow is greater than zero but cells are empty | C# code to identify Excel worksheets with no data despite a reported MaxDataRow | check for formatting‑only worksheets in a workbook using Aspose.Cells | determine if a worksheet has only styles and no values in .NET
// Tags: detect formatting‑only worksheets Aspose.Cells | maxdatarow without cell values .NET | worksheet data presence verification C# | scan worksheet cells for null values Aspose.Cells | flag empty data region Excel Aspose

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetFormattingOnlyFlagger
{
    // The example loads an Excel workbook, iterates through each worksheet, and uses Cells.MaxDataRow/MaxDataColumn to define the data region. It then scans every cell within that region; if MaxDataRow is greater than zero yet no cell contains a non‑null Value, the worksheet is flagged as containing only formatting.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Resolve input file
                string inputPath = args.Length > 0 ? args[0] : "input.xlsx";

                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: File \"{inputPath}\" does not exist.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Console.WriteLine($"Workbook \"{Path.GetFileName(inputPath)}\" loaded successfully.");
                Console.WriteLine($"Total worksheets: {workbook.Worksheets.Count}");

                bool anyFlagged = false;

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // MaxDataRow and MaxDataColumn are based on cells that contain *data*.
                    // Some UI‑oriented APIs (GridDesktop/GridWeb) may return a positive index even when only formatting exists.
                    // Therefore we perform an additional scan to verify the presence of actual values.
                    int maxDataRow = sheet.Cells.MaxDataRow;       // zero‑based index, -1 if no data
                    int maxDataColumn = sheet.Cells.MaxDataColumn; // zero‑based index, -1 if no data

                    // If MaxDataRow is -1, there is definitely no data – skip.
                    if (maxDataRow < 0)
                    {
                        continue;
                    }

                    // Flag condition: MaxDataRow > 0 (i.e., at least one row index beyond the header) 
                    // but *none* of the cells up to MaxDataRow/MaxDataColumn contain a value.
                    bool hasAnyValue = false;

                    // Scan the rectangular area defined by MaxDataRow/MaxDataColumn
                    for (int row = 0; row <= maxDataRow && !hasAnyValue; row++)
                    {
                        for (int col = 0; col <= maxDataColumn && !hasAnyValue; col++)
                        {
                            Cell cell = sheet.Cells[row, col];
                            // Cell.Value is null when the cell holds only formatting.
                            if (cell.Value != null)
                            {
                                hasAnyValue = true;
                            }
                        }
                    }

                    // If we have a MaxDataRow > 0 but no actual values, flag the worksheet.
                    if (maxDataRow > 0 && !hasAnyValue)
                    {
                        anyFlagged = true;
                        Console.WriteLine($"[FLAGGED] Worksheet \"{sheet.Name}\" (Index {sheet.Index}) appears to contain only formatting.");
                    }
                }

                if (!anyFlagged)
                {
                    Console.WriteLine("No worksheets with formatting‑only content were detected.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
