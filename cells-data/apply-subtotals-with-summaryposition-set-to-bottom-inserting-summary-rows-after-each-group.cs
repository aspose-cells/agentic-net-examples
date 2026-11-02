// Title: Insert subtotal rows after each grouped section in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add subtotal row after each group in Excel | How to programmatically insert a SUM formula for grouped data with Aspose.Cells | C# generate group totals and place a summary row at the bottom of each section using Aspose.Cells | Insert rows and set formulas after grouped rows in an Excel file with Aspose.Cells .NET | Automate Excel subtotal rows per category using Aspose.Cells library in C#
// Tags: Aspose.Cells insert rows programmatically | Aspose.Cells subtotal row automation | Aspose.Cells SUM formula insertion | Excel grouping totals C# | Aspose.Cells worksheet row insertion after group | C# Excel subtotal generation with Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, groups rows by the values in column A, inserts a new row after each group's last row, writes a "Subtotal" label in column A, adds a SUM formula in column C that totals the group's values, and saves the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output_with_subtotals.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                var workbook = new Workbook(inputPath);
                var worksheet = workbook.Worksheets[0];

                // Determine the used range of rows
                int firstDataRow = worksheet.Cells.MinDataRow;
                int lastDataRow = worksheet.Cells.MaxDataRow;

                // Column indexes (0‑based)
                int aggregateColumn = 2; // Column C (values to sum)
                int groupByColumn = 0;   // Column A (grouping column)

                // Collect rows for each group based on column A
                var groups = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);
                for (int row = firstDataRow; row <= lastDataRow; row++)
                {
                    var key = worksheet.Cells[row, groupByColumn].StringValue ?? string.Empty;
                    if (!groups.ContainsKey(key))
                        groups[key] = new List<int>();
                    groups[key].Add(row);
                }

                // Insert subtotal rows starting from the bottom to preserve indices
                foreach (var kvp in groups)
                {
                    var rows = kvp.Value;
                    if (rows.Count == 0) continue;

                    int startRow = rows[0];
                    int endRow = rows[rows.Count - 1];

                    // Insert a new row after the group's last row
                    int insertPos = endRow + 1;
                    worksheet.Cells.InsertRows(insertPos, 1);

                    // Write a label in the grouping column (optional)
                    worksheet.Cells[insertPos, groupByColumn].PutValue("Subtotal");

                    // Set a SUM formula in the aggregate column
                    // Excel rows are 1‑based, so add 1 to each index
                    string startCell = CellsHelper.CellIndexToName(startRow, aggregateColumn);
                    string endCell = CellsHelper.CellIndexToName(endRow, aggregateColumn);
                    string formula = $"=SUM({startCell}:{endCell})";
                    worksheet.Cells[insertPos, aggregateColumn].Formula = formula;
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
