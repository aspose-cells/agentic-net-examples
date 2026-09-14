// Title: Export Aspose.Cells formula evaluation order to a text file for debugging Excel dependencies (C#)
// AI Prompts: Write C# code that opens an .xlsx workbook with Aspose.Cells, calls CalculateFormula, and writes each formula cell's address, zero‑based row and column to a .txt file in the order they are evaluated. | Create a method that iterates through all worksheets in an Aspose.Cells Workbook, detects cells where IsFormula is true, and logs their name, row, column, and optionally the formula text to a plain‑text report. | Modify the exporter to include the evaluated value of each formula alongside its address and coordinates in the output file using Aspose.Cells APIs.
// Common Searches: how to get the order of formula calculation in Aspose.Cells C# | export list of formula cells with row and column indices from an Excel workbook using Aspose.Cells | debugging Excel formula dependencies with Aspose.Cells .NET example | write formula evaluation sequence to a text file using Aspose.Cells library
// Tags: aspocells export formula order txt | c# iterate worksheets formula cells | aspocells calculate workbook formulas | debug excel formula dependency aspocells | write cell address row column to file c#

using System;
using System.IO;
using Aspose.Cells;

// Loads 'input.xlsx' with Aspose.Cells, calculates all formulas, iterates through each worksheet and cell, and writes each formula cell's address, zero‑based row and column to 'FormulaEvaluationOrder.txt' in the evaluation order.
class FormulaEvaluationOrderExporter
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string workbookPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: The file '{workbookPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Prepare the output file path
            string outputPath = "FormulaEvaluationOrder.txt";

            using (StreamWriter writer = new StreamWriter(outputPath))
            {
                writer.WriteLine("Formula Evaluation Order:");
                writer.WriteLine("--------------------------");

                int index = 1;

                // Iterate through each worksheet and its cells to collect formula cells
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Loop through all used cells
                    foreach (Cell cell in cells)
                    {
                        if (cell.IsFormula)
                        {
                            // Cell.Name gives the address (e.g., "B2")
                            string cellName = cell.Name;
                            int row = cell.Row;       // zero‑based
                            int column = cell.Column; // zero‑based

                            writer.WriteLine($"{index}. Cell: {cellName} (Row: {row}, Column: {column})");
                            index++;
                        }
                    }
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
