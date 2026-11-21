// Title: Calculate a running total with Aspose.Cells Formula property while importing rows in C#
// Common Searches: C# Aspose.Cells set SUM formula for each row to compute cumulative total | How to add a running total column in Excel using Aspose.Cells during data import | Aspose.Cells Formula property example for cumulative sum in a worksheet | Create running total column programmatically with Aspose.Cells C# | Update Excel cells with dynamic SUM formula while looping through data in C#
// Tags: Aspose.Cells set cell formula C# | Workbook.CalculateFormula method C# | Excel SUM formula generation Aspose.Cells | Data import loop Aspose.Cells | Formula property usage C#

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace RunningTotalExample
{
    // The example creates a new workbook, adds headers, iterates over a list of amounts, writes each amount to column A, assigns a SUM formula in column B that references all previous rows to produce a running total, forces formula recalculation, and saves the file as RunningTotal.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data to be imported row by row
            List<double> data = new List<double> { 120.5, 85.0, 150.75, 60.25, 200.0 };

            // Header row
            cells["A1"].PutValue("Amount");
            cells["B1"].PutValue("Running Total");

            // Import data and set running total formula for each row
            for (int i = 0; i < data.Count; i++)
            {
                int rowIndex = i + 1; // +1 because row 0 is the header

                // Insert the amount value in column A
                cells[rowIndex, 0].PutValue(data[i]);

                // Set the running total formula in column B
                // Formula: =SUM($A$2:A{currentRow})
                // Excel uses 1‑based indexing, so we add 1 to rowIndex for the cell address
                string formula = $"SUM($A$2:A{rowIndex + 1})";
                cells[rowIndex, 1].Formula = formula;
            }

            // Recalculate all formulas so that running totals are materialized
            workbook.CalculateFormula();

            // Save the workbook (lifecycle rule)
            workbook.Save("RunningTotal.xlsx");
        }
    }
}
