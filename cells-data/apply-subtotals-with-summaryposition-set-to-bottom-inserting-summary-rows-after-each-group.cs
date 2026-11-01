// Title: Add bottom‑positioned subtotal rows for each category group in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: how to add a subtotal row after each category group using Aspose.Cells in C# | C# Aspose.Cells insert summary rows at the bottom of each group | calculate group totals and insert rows without affecting other rows with Aspose.Cells | Aspose.Cells set summaryPosition to bottom for grouped data | manual subtotal implementation example with Aspose.Cells for .NET
// Tags: Aspose.Cells group subtotal calculation | C# insert rows without shifting data Aspose.Cells | Excel bottom summary position Aspose.Cells | manual subtotal implementation Aspose.Cells | Aspose.Cells aggregate column values by group

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads or creates an Excel workbook, groups rows by the Category column, computes the sum of the Amount column for each group, and inserts a subtotal row directly after each group at the bottom using Aspose.Cells for .NET before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook or create a new one if the file is missing
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                // Sample data to demonstrate subtotal functionality
                ws.Cells["A1"].PutValue("Category");
                ws.Cells["B1"].PutValue("Item");
                ws.Cells["C1"].PutValue("Amount");
                ws.Cells["A2"].PutValue("A");
                ws.Cells["B2"].PutValue("Item1");
                ws.Cells["C2"].PutValue(10);
                ws.Cells["A3"].PutValue("A");
                ws.Cells["B3"].PutValue("Item2");
                ws.Cells["C3"].PutValue(20);
                ws.Cells["A4"].PutValue("B");
                ws.Cells["B4"].PutValue("Item3");
                ws.Cells["C4"].PutValue(15);
                ws.Cells["A5"].PutValue("B");
                ws.Cells["B5"].PutValue("Item4");
                ws.Cells["C5"].PutValue(25);
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Determine the data range (including header row)
            int startRow = 0;          // zero‑based first row (header)
            int startColumn = 0;       // zero‑based first column (A)
            int lastDataRow = sheet.Cells.MaxDataRow;
            int lastDataColumn = sheet.Cells.MaxDataColumn;

            // Manual subtotal implementation (group by column A, sum column C)
            var subtotals = new List<(int RowIndex, double Sum)>();
            string previousCategory = null;
            double groupSum = 0;

            // Iterate through data rows (skip header)
            for (int row = startRow + 1; row <= lastDataRow; row++)
            {
                string category = sheet.Cells[row, startColumn].StringValue;
                double amount = sheet.Cells[row, startColumn + 2].DoubleValue; // column C

                if (previousCategory == null)
                {
                    previousCategory = category;
                }

                if (category != previousCategory)
                {
                    // End of previous group – store subtotal position
                    subtotals.Add((row - 1, groupSum));
                    // Reset for new group
                    previousCategory = category;
                    groupSum = 0;
                }

                groupSum += amount;
            }

            // Add subtotal for the final group
            if (previousCategory != null)
            {
                subtotals.Add((lastDataRow, groupSum));
            }

            // Insert subtotal rows from bottom to top to avoid index shifting
            for (int i = subtotals.Count - 1; i >= 0; i--)
            {
                int insertAfter = subtotals[i].RowIndex;
                double sum = subtotals[i].Sum;

                sheet.Cells.InsertRows(insertAfter + 1, 1);
                sheet.Cells[insertAfter + 1, startColumn + 1].PutValue("Subtotal"); // column B
                sheet.Cells[insertAfter + 1, startColumn + 2].PutValue(sum);      // column C
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
