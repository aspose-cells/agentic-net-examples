// Title: Insert MAX subtotal rows after each group in column A, summarizing column C with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to detect changes in column A and insert a row after each group that contains a MAX formula for the corresponding range in column C. | Modify the provided Aspose.Cells example to place the subtotal row before each group while still calculating the MAX of column C for that group. | Adapt the solution to compute SUM instead of MAX for column C subtotals, keeping the same group‑by‑column‑A logic.
// Common Searches: aspocells c# add max subtotal after each category in excel column a | how to insert subtotal rows based on group changes using Aspose.Cells .NET | c# generate dynamic max formula for column C groups with Aspose.Cells | aspocells insert rows programmatically after grouping column values | excel subtotal max per group using Aspose.Cells C# example
// Tags: Aspose.Cells aggregate row creation | aggregate calculation with Aspose.Cells | C# detect column value changes in Excel | dynamic range address generation .NET | Excel workbook grouping automation

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The C# program loads an Excel workbook, scans column A for group transitions, and inserts a new row after each group containing a MAX formula that aggregates the values in column C for that group. It also adds a final subtotal for the last group and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet
                Worksheet ws = workbook.Worksheets[0];

                // Column indices (0‑based): A = 0 (group), C = 2 (values)
                int groupCol = 0;
                int valueCol = 2;

                // Determine the used range
                int firstDataRow = ws.Cells.MinDataRow;
                int lastDataRow = ws.Cells.MaxDataRow;

                // Initialize tracking for the first group
                int groupStartRow = firstDataRow;
                string currentGroup = ws.Cells[groupStartRow, groupCol].StringValue;

                // Iterate through rows to detect group changes
                for (int row = firstDataRow; row <= lastDataRow; row++)
                {
                    string cellValue = ws.Cells[row, groupCol].StringValue;

                    // When the group value changes, insert a subtotal row for the previous group
                    if (cellValue != currentGroup)
                    {
                        int insertPos = row;

                        // Insert a new row for the subtotal
                        ws.Cells.InsertRows(insertPos, 1);
                        ws.Cells[insertPos, groupCol].PutValue($"{currentGroup} Max");

                        // Build the address range for the MAX formula (C column of the group)
                        string startAddr = ws.Cells[groupStartRow, valueCol].Name;
                        string endAddr = ws.Cells[row - 1, valueCol].Name;
                        ws.Cells[insertPos, valueCol].Formula = $"MAX({startAddr}:{endAddr})";

                        // Update tracking variables for the new group
                        currentGroup = cellValue;
                        groupStartRow = row + 1; // +1 because we just inserted a row

                        // Adjust loop counters because the sheet grew by one row
                        lastDataRow++;
                        row++; // skip the newly inserted subtotal row
                    }
                }

                // After the loop, handle the final group
                int finalInsertPos = lastDataRow + 1;
                ws.Cells.InsertRows(finalInsertPos, 1);
                ws.Cells[finalInsertPos, groupCol].PutValue($"{currentGroup} Max");
                string startAddrFinal = ws.Cells[groupStartRow, valueCol].Name;
                string endAddrFinal = ws.Cells[lastDataRow, valueCol].Name;
                ws.Cells[finalInsertPos, valueCol].Formula = $"MAX({startAddrFinal}:{endAddrFinal})";

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
