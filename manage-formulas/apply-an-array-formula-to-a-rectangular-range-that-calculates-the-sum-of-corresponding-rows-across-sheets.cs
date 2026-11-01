// Title: Generate a summary worksheet with row‑wise SUM formulas that reference multiple sheets using Aspose.Cells for .NET (C#)
// Common Searches: how to create a summary sheet that sums rows from multiple worksheets using Aspose.Cells C# | Aspose.Cells C# set formula =SUM(Sheet1!A1,Sheet2!A1,Sheet3!A1) for a column | C# code to add cross‑sheet SUM formulas to a new worksheet with Aspose.Cells | populate a column with formulas that total corresponding rows across sheets in Aspose.Cells | save workbook with cross‑sheet formulas as .xlsx using Aspose.Cells .NET
// Tags: Aspose.Cells cross-sheet SUM formula C# | generate summary worksheet with row totals Aspose.Cells | populate column A with formulas referencing multiple sheets | save workbook as XLSX using Aspose.Cells .NET | programmatic worksheet creation and formula assignment Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, adds three worksheets (Sheet1‑Sheet3) with numeric data in column A, then adds a Summary worksheet where each cell in column A contains a formula like =SUM(Sheet1!A1,Sheet2!A1,Sheet3!A1) to total the corresponding rows across the three sheets, and finally saves the file as ArrayFormulaAcrossSheets.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add three worksheets with sample numeric data
            for (int i = 0; i < 3; i++)
            {
                Worksheet sheet;
                // Use existing sheet if present, otherwise add a new one
                if (i < workbook.Worksheets.Count)
                {
                    sheet = workbook.Worksheets[i];
                }
                else
                {
                    // Worksheets.Add() returns the index of the new sheet in some versions,
                    // so retrieve the sheet by that index.
                    int newIndex = workbook.Worksheets.Add();
                    sheet = workbook.Worksheets[newIndex];
                }

                sheet.Name = $"Sheet{i + 1}";

                // Populate cells A1:A5 with values (i+1) * rowIndex
                for (int row = 0; row < 5; row++)
                {
                    sheet.Cells[row, 0].PutValue((i + 1) * (row + 1));
                }
            }

            // Add a summary worksheet where the formulas will be placed
            Worksheet summary;
            {
                int summaryIndex = workbook.Worksheets.Add();
                summary = workbook.Worksheets[summaryIndex];
            }
            summary.Name = "Summary";

            // Apply formulas to each cell in A1:A5 that sum the corresponding rows across the three sheets
            for (int row = 0; row < 5; row++)
            {
                // Example: =SUM(Sheet1!A1,Sheet2!A1,Sheet3!A1)
                summary.Cells[row, 0].Formula = $"=SUM(Sheet1!A{row + 1},Sheet2!A{row + 1},Sheet3!A{row + 1})";
            }

            // Determine output path and ensure directory exists
            string outputFile = "ArrayFormulaAcrossSheets.xlsx";
            string outputPath = Path.GetFullPath(outputFile);
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
