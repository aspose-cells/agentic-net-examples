// Title: Add a SUM(E2:E100) formula to cell E101 and save the workbook with Aspose.Cells for .NET (C#)
// Description: Creates a new Workbook, sets the formula "=SUM(E2:E100)" in cell E101, forces calculation of all formulas, and saves the file as SumColumnE.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# SUM formula | add SUM to Excel with Aspose.Cells | set cell formula Aspose.Cells .NET | calculate workbook formulas C# | save Excel file Aspose.Cells
// Common Searches: Aspose.Cells how to insert SUM formula in C# | C# set formula E101 to sum E2:E100 Aspose.Cells | calculate and save workbook after adding formula Aspose.Cells .NET | example code for SUM(E2:E100) using Aspose.Cells
// Developer Intent: Insert a SUM formula that totals cells E2 through E100, recalculate the workbook, and write the result to a file.
// Use Cases: Automatically total expense entries in a financial report footer. | Provide a live column total in a data‑entry template that updates as rows are filled. | Consolidate monthly sales figures by summing a column range before exporting the workbook.
// AI Prompts: Generate C# code with Aspose.Cells that places "=SUM(E2:E100)" in E101, recalculates, and saves the workbook. | Show how to add dynamic SUM formulas for multiple columns (e.g., E, F, G) using Aspose.Cells for .NET. | Explain how to detect the last used row in column E and set a range‑aware SUM formula with Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a new Workbook, sets the formula "=SUM(E2:E100)" in cell E101, forces calculation of all formulas, and saves the file as SumColumnE.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (create rule)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Place the sum formula in a cell (e.g., E101) to sum E2:E100
        // Formula strings must start with '=' and use commas as delimiters if needed
        worksheet.Cells["E101"].Formula = "=SUM(E2:E100)";

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Save the workbook to a file (save rule)
        workbook.Save("SumColumnE.xlsx");
    }
}
