// Title: Detect and list cells that return #DIV/0! errors after calculating formulas with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to calculate all formulas in a workbook and return a List<string> of cell addresses where the result is a #DIV/0! error. | Create a reusable method that scans every worksheet’s used range, checks CellValueType.IsError, and collects the names of cells that contain any formula error. | Extend the solution to separate cells by error type (e.g., #DIV/0!, #VALUE!, #REF!) and output each group as a distinct list.
// Common Searches: Aspose.Cells C# find cells with #DIV/0! after CalculateFormula | How to enumerate Excel formula errors using Aspose.Cells .NET | Get addresses of division by zero errors in a workbook with Aspose.Cells | C# iterate used range to collect error cells in Aspose.Cells workbook | Programmatically list cells that contain #DIV/0! in Excel using Aspose.Cells
// Tags: Aspose.Cells calculate formulas and detect division by zero errors | C# enumerate error cells in Excel workbook | retrieve cell addresses of #DIV/0! using Aspose.Cells | iterate used range to collect formula error values | Workbook error handling for #DIV/0! in .NET

using Aspose.Cells;
using System;
using System.Collections.Generic;

// The example creates a workbook, inserts values and formulas that may cause division by zero, runs CalculateFormula, iterates each worksheet's used range, checks for CellValueType.IsError, gathers the addresses of cells that evaluate to #DIV/0! (or any error), prints the list, and saves the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and access the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate cells with values and formulas that may cause division by zero
            ws.Cells["A1"].PutValue(10);      // Numerator
            ws.Cells["A2"].PutValue(0);       // Denominator (zero)
            ws.Cells["B1"].Formula = "A1/A2"; // Will result in #DIV/0!
            ws.Cells["B2"].Formula = "A1/5";  // Normal division

            // Calculate all formulas in the workbook
            wb.CalculateFormula();

            // Collect addresses of cells that contain #DIV/0! errors
            List<string> divZeroCells = new List<string>();

            foreach (Worksheet sheet in wb.Worksheets)
            {
                // Get the used range of the sheet
                var usedRange = sheet.Cells.MaxDisplayRange;
                int startRow = usedRange.FirstRow;
                int endRow = startRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = startCol + usedRange.ColumnCount - 1;

                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // If the cell contains any error (including #DIV/0!), record its address
                        if (cell.Type == CellValueType.IsError)
                        {
                            divZeroCells.Add(cell.Name);
                        }
                    }
                }
            }

            // Output the collected cell addresses
            Console.WriteLine("Cells with #DIV/0! errors:");
            foreach (string address in divZeroCells)
            {
                Console.WriteLine(address);
            }

            // Save the workbook (optional)
            wb.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
