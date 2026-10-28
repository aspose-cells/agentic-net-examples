// Title: Apply row‑wise SUM formulas to a data table using Aspose.Cells for .NET (C#)
// Common Searches: how to add a SUM formula to each row in an Excel sheet with Aspose.Cells C# | Aspose.Cells C# set formula for a range of cells programmatically | calculate row totals in a worksheet without using array formulas in Aspose.Cells | save an Excel workbook after evaluating formulas with Aspose.Cells .NET | populate sample data and apply formulas using Aspose.Cells for .NET
// Tags: row-wise sum formula Aspose.Cells C# | set cell formula programmatically Aspose.Cells | calculate workbook formulas .NET | populate Excel data Aspose.Cells | save workbook after formula evaluation Aspose.Cells | apply SUM across columns Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, fills rows 2‑10 with numeric data, assigns a standard SUM formula to each cell in column E to total columns A‑D for that row, forces formula calculation, and saves the workbook as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Populate sample data in rows 2-10 (A2:D10)
            for (int i = 0; i < 9; i++) // 9 rows
            {
                sheet.Cells[i + 1, 0].PutValue(i + 1);           // A column
                sheet.Cells[i + 1, 1].PutValue((i + 1) * 2);    // B column
                sheet.Cells[i + 1, 2].PutValue((i + 1) * 3);    // C column
                sheet.Cells[i + 1, 3].PutValue((i + 1) * 4);    // D column
            }

            // Apply a row‑wise sum formula in column E (E2:E10)
            for (int i = 0; i < 9; i++)
            {
                // Use a normal formula for each cell (no array formula needed)
                sheet.Cells[i + 1, 4].Formula = $"=SUM(A{i + 2}:D{i + 2})";
            }

            // Force calculation of all formulas
            workbook.CalculateFormula();

            // Determine output file path (ensure absolute path)
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ArrayFormulaResult.xlsx");
            string outputDir = Path.GetDirectoryName(outputPath) ?? Directory.GetCurrentDirectory();

            // Ensure the output directory exists
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
