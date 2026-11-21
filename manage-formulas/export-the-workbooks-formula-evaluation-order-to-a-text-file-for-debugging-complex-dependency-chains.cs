// Title: Export the order of formula cells from an Aspose.Cells workbook to a text file using C#
// Common Searches: C# Aspose.Cells export list of formula cell addresses to txt | How to get formula evaluation sequence from an Excel file with Aspose.Cells | Debugging Excel formula dependencies by writing cell order to a file in C# | Retrieve formula cells in worksheet order using Aspose.Cells API | Save Aspose.Cells formula cell addresses for troubleshooting complex formulas
// Tags: Aspose.Cells export formula addresses to text | C# retrieve formula cell order from workbook | debug Excel formula dependency chain Aspose | write formula cell list using Aspose.Cells | iterate worksheets for formulas Aspose.Cells

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, scans each worksheet in row‑major order, collects the A1‑style addresses of all formula cells (including the sheet name), and writes these addresses to a text file (FormulaEvaluationOrder.txt). This output helps developers debug complex formula dependency chains by revealing the approximate evaluation order.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "FormulaEvaluationOrder.txt";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Collect formula cells in evaluation order (approximate by worksheet order then row/column)
            List<string> formulaEntries = new List<string>();

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the maximum used row and column to limit the scan
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        if (cell != null && cell.IsFormula)
                        {
                            // Convert row/column to A1 style address
                            string cellAddress = CellsHelper.CellIndexToName(row, col);
                            formulaEntries.Add($"{sheet.Name}!{cellAddress}");
                        }
                    }
                }
            }

            // Write the collected formula addresses to the output file
            using (StreamWriter writer = new StreamWriter(outputPath))
            {
                foreach (string entry in formulaEntries)
                {
                    writer.WriteLine(entry);
                }
            }

            Console.WriteLine($"Formula evaluation order has been exported to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
