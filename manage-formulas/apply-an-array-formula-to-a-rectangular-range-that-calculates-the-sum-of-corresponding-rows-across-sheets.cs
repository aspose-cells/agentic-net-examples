// Title: Create a summary worksheet with row‑by‑row cross‑sheet sum formulas using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that builds three worksheets, fills column A with sample numbers, adds a fourth worksheet named Summary, and writes formulas in cells A2:A5 that add the values from the same rows of the three source sheets. | Show how to programmatically insert a cross‑sheet sum formula into a rectangular range and save the workbook as an .xlsx file using Aspose.Cells for .NET.
// Common Searches: asp.net aspose.cells add summary sheet that sums rows from multiple worksheets | c# generate cross-sheet sum formulas in a range with Aspose.Cells | how to write formulas that reference other sheets in Aspose.Cells C# example | apply array formula across rows to sum values from several worksheets using Aspose.Cells | save workbook with cross-sheet formulas as xlsx using Aspose.Cells .NET
// Tags: cross-sheet sum formula Aspose.Cells | populate summary worksheet C# Aspose.Cells | write array formula across worksheets | save workbook as xlsx Aspose.Cells | add multiple worksheets programmatically Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsArrayFormulaExample
{
    // The example creates a new workbook, adds three data worksheets (Sheet1, Sheet2, Sheet3) and fills A2:A5 on each with incremental numbers. A fourth worksheet named Summary is added, and cells A2:A5 on this sheet receive formulas that sum the corresponding rows from the three source sheets (e.g., Sheet1!A2+Sheet2!A2+Sheet3!A2). The workbook is saved as 'ArrayFormula_SumAcrossSheets.xlsx'.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add three worksheets that will contain the source data
                Worksheet sheet1 = workbook.Worksheets[workbook.Worksheets.Add()];
                sheet1.Name = "Sheet1";
                Worksheet sheet2 = workbook.Worksheets[workbook.Worksheets.Add()];
                sheet2.Name = "Sheet2";
                Worksheet sheet3 = workbook.Worksheets[workbook.Worksheets.Add()];
                sheet3.Name = "Sheet3";

                // Fill each source sheet with sample numeric data (rows 2‑5, column A)
                for (int row = 1; row <= 4; row++) // zero‑based index, so row 1 = Excel row 2
                {
                    sheet1.Cells[row, 0].PutValue(row * 1);   // Sheet1: 1,2,3,4
                    sheet2.Cells[row, 0].PutValue(row * 10);  // Sheet2: 10,20,30,40
                    sheet3.Cells[row, 0].PutValue(row * 100); // Sheet3: 100,200,300,400
                }

                // Add a summary worksheet where the results will be placed
                Worksheet summary = workbook.Worksheets[workbook.Worksheets.Add()];
                summary.Name = "Summary";

                // Populate the summary sheet with formulas that sum the corresponding rows across the three sheets
                // (A2:A5)
                int startRow = 1; // zero‑based index for Excel row 2
                for (int i = 0; i < 4; i++)
                {
                    int excelRow = startRow + i + 1; // Excel row number (2‑5)
                    Cell cell = summary.Cells[startRow + i, 0];
                    // Simple sum formula for the current row across the three sheets
                    cell.Formula = $"Sheet1!A{excelRow}+Sheet2!A{excelRow}+Sheet3!A{excelRow}";
                }

                // Save the workbook to a file
                string outputPath = "ArrayFormula_SumAcrossSheets.xlsx";

                // Ensure the directory exists before saving
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
