// Title: Find and list all #DIV/0! error cells in an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file using Aspose.Cells, forces a full formula recalculation, and returns a List<Cell> containing every cell whose value is the #DIV/0! error. | Extend the previous example to apply a red background style to each #DIV/0! cell while still outputting their addresses and worksheet names.
// Common Searches: aspocells c# find cells containing #DIV/0! error | how to list division by zero error cells in Excel using Aspose.Cells .NET | enumerate formula error values in a workbook with Aspose.Cells | retrieve addresses of #DIV/0! cells from used range Aspose.Cells C# | scan worksheets for Excel error values with Aspose.Cells
// Tags: detect #DIV/0! errors Aspose.Cells | collect cells with division by zero error .NET | enumerate formula error cells Excel Aspose | scan worksheets used range for error values | highlight #DIV/0! cells Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The code loads a workbook, recalculates all formulas, iterates each worksheet's used range, gathers cells whose value equals "#DIV/0!", and prints their addresses together with the worksheet name. Optional styling can be added to highlight those cells.
class DivZeroFinder
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the workbook
                workbook = new Workbook(inputPath);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                return;
            }

            // Calculate all formulas so error values are up‑to‑date
            try
            {
                workbook.CalculateFormula();
            }
            catch (Exception calcEx)
            {
                Console.WriteLine($"Formula calculation error: {calcEx.Message}");
            }

            // List to hold cells that contain #DIV/0! errors
            List<Cell> divZeroCells = new List<Cell>();

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range of the worksheet
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                // Iterate through each cell in the used range
                foreach (Cell cell in usedRange)
                {
                    // Check if the cell contains an error value and if it is #DIV/0!
                    if (cell.Type == CellValueType.IsError && cell.StringValue == "#DIV/0!")
                    {
                        divZeroCells.Add(cell);
                    }
                }
            }

            // Output the addresses of the cells with #DIV/0! errors
            Console.WriteLine("Cells containing #DIV/0! errors:");
            foreach (Cell errCell in divZeroCells)
            {
                Console.WriteLine($"{errCell.Name} (Worksheet: {errCell.Worksheet.Name})");
            }

            // (Optional) Save the workbook if you made any changes
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
