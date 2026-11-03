// Title: Generate a summary worksheet that counts Excel cell data types using Aspose.Cells for .NET (C#)
// Common Searches: how to count different cell value types in an Excel workbook using Aspose.Cells C# | list all populated cells and calculate CellValueType distribution with Aspose.Cells | create a summary sheet with data type counts in Aspose.Cells for .NET | C# code to generate a data type report from an Excel file with Aspose.Cells | auto-fit columns after writing a summary worksheet with Aspose.Cells C#
// Tags: cell value type counting Aspose.Cells | data type count worksheet generation | enumerate used cells Aspose.Cells C# | write CellValueType frequency to Excel | adjust column width Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDataTypeSummary
{
    // The program loads an existing Excel workbook, iterates over every used cell in the first worksheet, tallies each CellValueType into a dictionary, adds a new worksheet named "Summary", writes the data type names and their non‑zero counts, adjusts column widths for readability, and saves the updated workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the source workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet
                Worksheet sourceSheet = workbook.Worksheets[0];
                Cells cells = sourceSheet.Cells;

                // Prepare dictionary for counting each CellValueType
                Dictionary<CellValueType, int> typeCounts = new Dictionary<CellValueType, int>();
                foreach (CellValueType type in Enum.GetValues(typeof(CellValueType)))
                {
                    typeCounts[type] = 0;
                }

                // Enumerate used cells
                int maxRow = cells.MaxDataRow;
                int maxColumn = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell == null) continue;

                        // Use the correct property to get the cell's value type
                        CellValueType valueType = cell.Type;
                        typeCounts[valueType]++;
                    }
                }

                // Add a new worksheet for the summary report
                Worksheet summarySheet = workbook.Worksheets.Add("Summary");
                Cells summaryCells = summarySheet.Cells;

                // Write header
                summaryCells[0, 0].PutValue("Data Type");
                summaryCells[0, 1].PutValue("Count");

                // Write counts (only types with count > 0)
                int summaryRow = 1;
                foreach (var kvp in typeCounts)
                {
                    if (kvp.Value > 0)
                    {
                        summaryCells[summaryRow, 0].PutValue(kvp.Key.ToString());
                        summaryCells[summaryRow, 1].PutValue(kvp.Value);
                        summaryRow++;
                    }
                }

                // Auto-fit columns for readability
                summarySheet.AutoFitColumns();

                // Save the workbook with the summary
                workbook.Save(outputPath);
                Console.WriteLine($"Summary saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
