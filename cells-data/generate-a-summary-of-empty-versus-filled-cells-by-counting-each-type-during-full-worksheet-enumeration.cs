// Title: C# Aspose.Cells – Count Empty vs Filled Cells with CellValueType Breakdown
// Description: Iterate a worksheet using MaxDataRow/Column and CheckCell, count empty and filled cells, group filled cells by CellValueType, display totals and save the workbook.
// Keywords: Aspose.Cells empty cells count | C# cell type summary | CellValueType statistics | enumerate worksheet cells .NET | CountLarge vs actual cell count | Excel data completeness Aspose | Aspose.Cells example C#
// Common Searches: how to count empty cells in Aspose.Cells | cell type distribution Aspose.Cells C# | MaxDataRow iteration Aspose.Cells | Aspose.Cells CountLarge meaning | C# enumerate all worksheet cells Aspose
// Developer Intent: Enumerate every cell in a worksheet and obtain totals for empty and filled cells, plus a per‑CellValueType count of the filled cells.
// Use Cases: Create a data‑completeness report before publishing a workbook. | Detect missing values for automated data‑cleaning pipelines. | Profile memory usage by analyzing the distribution of cell types in large Excel files.
// AI Prompts: Generate C# Aspose.Cells code that scans a worksheet, returns a dictionary of CellValueType counts, and provides total empty and filled cell numbers. | Explain how to handle a completely empty worksheet when using MaxDataRow and MaxDataColumn with Aspose.Cells. | Suggest an efficient method to count only instantiated cells without iterating over every row and column in a massive sheet.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Iterate a worksheet using MaxDataRow/Column and CheckCell, count empty and filled cells, group filled cells by CellValueType, display totals and save the workbook.
class EmptyVsFilledSummary
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data (replace with loading an existing file if needed)
            cells["A1"].PutValue("Hello");
            cells["B1"].PutValue(123);
            cells["C1"].PutValue(DateTime.Now);
            cells["D1"].PutValue(true);
            // E1 left empty on purpose

            // Dictionaries to hold counts per CellValueType
            Dictionary<CellValueType, long> typeCounts = new Dictionary<CellValueType, long>();
            long emptyCount = 0;
            long filledCount = 0;

            // Determine the area that may contain data
            int maxRow = cells.MaxDataRow;       // last row that has data
            int maxCol = cells.MaxDataColumn;    // last column that has data

            // If the sheet is completely empty, MaxDataRow/Column are -1.
            // In that case we still want to report zero cells.
            if (maxRow >= 0 && maxCol >= 0)
            {
                for (int r = 0; r <= maxRow; r++)
                {
                    for (int c = 0; c <= maxCol; c++)
                    {
                        // CheckCell returns null if the cell has never been instantiated.
                        Cell cell = cells.CheckCell(r, c);
                        if (cell == null)
                        {
                            emptyCount++;
                            continue;
                        }

                        CellValueType ct = cell.Type;
                        if (ct == CellValueType.IsNull)
                        {
                            emptyCount++;
                        }
                        else
                        {
                            filledCount++;
                            if (!typeCounts.ContainsKey(ct))
                                typeCounts[ct] = 0;
                            typeCounts[ct]++;
                        }
                    }
                }
            }

            // Output the summary
            Console.WriteLine($"Total instantiated cells (CountLarge): {cells.CountLarge}");
            Console.WriteLine($"Empty cells: {emptyCount}");
            Console.WriteLine($"Filled cells: {filledCount}");
            Console.WriteLine("Filled cells by type:");
            foreach (var kvp in typeCounts)
            {
                Console.WriteLine($"  {kvp.Key}: {kvp.Value}");
            }

            // Save the workbook (optional)
            string outputPath = "EmptyVsFilledSummary.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
