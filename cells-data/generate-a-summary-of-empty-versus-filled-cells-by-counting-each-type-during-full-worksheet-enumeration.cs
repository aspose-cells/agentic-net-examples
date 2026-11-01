// Title: How to count empty cells and report filled cell types across an entire worksheet using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells count empty cells and list cell value types in a worksheet | enumerate all cells in an Excel file with Aspose.Cells and get statistics for each CellValueType | Aspose.Cells .NET get empty cell count and distribution of data types in used range | how to summarize filled cell types by CellValueType while iterating worksheet cells in C# | determine total, empty and typed cell counts in an XLSX using Aspose.Cells API
// Tags: Aspose.Cells empty cell tally .NET | CellValueType frequency analysis Aspose.Cells | worksheet used range iteration Aspose.Cells | Excel cell data summary C# | Aspose.Cells cell enumeration overview

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an XLSX workbook with Aspose.Cells, iterates each worksheet, determines the used range via MaxDataRow/MaxDataColumn, and walks every cell in that range. It counts total cells, increments a counter for cells whose Type is IsNull (empty), and aggregates non‑empty cells by their CellValueType in a dictionary. After processing a sheet it prints the worksheet name, total cells examined, empty cell count, and the count for each filled cell type, then saves the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        try
        {
            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int totalCells = 0;
                int emptyCells = 0;
                var filledByType = new Dictionary<CellValueType, int>();

                // Determine used range; MaxDataRow/Column are -1 for empty sheets
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                if (maxRow < 0 || maxCol < 0)
                {
                    Console.WriteLine($"Worksheet \"{sheet.Name}\" contains no data.");
                    continue;
                }

                // Enumerate cells within the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        totalCells++;

                        // In Aspose.Cells, a cell is considered empty when its type is IsNull
                        if (cell.Type == CellValueType.IsNull)
                        {
                            emptyCells++;
                        }
                        else
                        {
                            CellValueType type = cell.Type;
                            if (!filledByType.ContainsKey(type))
                                filledByType[type] = 0;
                            filledByType[type]++;
                        }
                    }
                }

                // Output summary for the worksheet
                Console.WriteLine($"Worksheet: {sheet.Name}");
                Console.WriteLine($"Total cells examined: {totalCells}");
                Console.WriteLine($"Empty cells: {emptyCells}");
                foreach (var kvp in filledByType)
                {
                    Console.WriteLine($"{kvp.Key}: {kvp.Value}");
                }
                Console.WriteLine();
            }

            // Save workbook (no modifications made, but required by lifecycle rules)
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during processing: {ex.Message}");
        }
    }
}
