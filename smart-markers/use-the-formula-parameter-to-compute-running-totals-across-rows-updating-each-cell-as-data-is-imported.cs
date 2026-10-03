// Title: Generate a running total column with a cumulative SUM formula while importing rows using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that imports a list of numbers into column A and assigns a cumulative SUM formula to column B for each row using Aspose.Cells. | Adjust the formula so the running total always starts from the first data row, regardless of the header position. | Add steps to recalculate all formulas after data import and save the workbook as an Excel file.
// Common Searches: Aspose.Cells C# how to create a cumulative sum column while importing data row by row | set running total formula with absolute start reference in Aspose.Cells workbook | calculate running totals in Excel using Aspose.Cells Formula property in .NET | dynamic SUM range for each row using Aspose.Cells smart markers | C# example of cumulative total column generation with Aspose.Cells
// Tags: set cell formula Aspose.Cells C# | Excel column cumulative total generation | import numeric list Aspose.Cells smart markers | recalculate workbook formulas Aspose.Cells | dynamic SUM range per row Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example creates a new workbook, adds headers, imports a list of numeric values into column A, assigns a cumulative SUM formula to column B using an absolute start reference (A$2) and a relative end row, recalculates all formulas, and saves the file as RunningTotal.xlsx.
class RunningTotalExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample data to import (could be from any source)
        List<double> values = new List<double> { 10, 20, 15, 30, 25 };

        // Header row
        cells["A1"].PutValue("Value");
        cells["B1"].PutValue("Running Total");

        // Import data row by row and set running total formula
        for (int i = 0; i < values.Count; i++)
        {
            int rowIndex = i + 1; // +1 because row 0 is the header

            // Import the value into column A
            cells[rowIndex, 0].PutValue(values[i]);

            // Set the running total formula in column B
            // Formula: =SUM(A$2:A{currentRow})
            // Using absolute reference for the start row (A$2) and relative end row
            string formula = $"=SUM(A$2:A{rowIndex + 1})";
            cells[rowIndex, 1].Formula = formula;
        }

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Save the workbook to a file
        workbook.Save("RunningTotal.xlsx");
    }
}
