// Title: How to insert a SUM formula in cell G10 that totals B2:B9 using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add SUM formula to cell G10 for range B2:B9 | set formula =SUM(B2:B9) in worksheet using Aspose.Cells .NET | calculate total of column B rows 2 to 9 and write result to G10 with Aspose.Cells | evaluate all formulas after inserting SUM in an Aspose.Cells workbook C# | save workbook as XLSX after adding a formula with Aspose.Cells
// Tags: Aspose.Cells set cell formula C# | Aspose.Cells calculate workbook formulas | Aspose.Cells save workbook as Xlsx | Aspose.Cells populate numeric cells | Aspose.Cells worksheet formula evaluation

using System;
using Aspose.Cells;

// The example creates a new XLSX workbook, fills cells B2 through B9 with numeric values, assigns the formula "=SUM(B2:B9)" to cell G10, forces calculation of all formulas, saves the file as FormulaDemo.xlsx, and prints both the formula text and its computed result.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (default format is XLSX)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate cells B2:B9 with sample numeric data (optional, for demonstration)
            // B column has index 1 (zero‑based). Rows 2‑9 correspond to indices 1‑8.
            for (int rowIndex = 1; rowIndex <= 8; rowIndex++)
            {
                // Example value: 10, 20, 30, ... 80
                worksheet.Cells[rowIndex, 1].PutValue(rowIndex * 10);
            }

            // Insert the SUM formula into cell G10.
            // G column has index 6, row 10 has index 9 (zero‑based).
            // Using the address notation is clearer.
            worksheet.Cells["G10"].Formula = "=SUM(B2:B9)";

            // Calculate all formulas in the workbook so that G10 contains the result.
            workbook.CalculateFormula();

            // Save the workbook to a file in the current directory.
            string outputFile = "FormulaDemo.xlsx";
            workbook.Save(outputFile, SaveFormat.Xlsx);

            // Output confirmation and the computed value.
            Console.WriteLine($"Workbook saved as \"{outputFile}\".");
            Console.WriteLine($"Formula in G10: {worksheet.Cells["G10"].Formula}");
            Console.WriteLine($"Calculated value in G10: {worksheet.Cells["G10"].Value}");
        }
        catch (Exception ex)
        {
            // Any unexpected errors are reported to the console.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
