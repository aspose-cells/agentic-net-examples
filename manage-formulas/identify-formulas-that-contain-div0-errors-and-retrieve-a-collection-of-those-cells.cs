// Title: Locate #DIV/0! formula errors in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# find cells that return #DIV/0! error | C# iterate worksheets to collect division by zero formula errors using Aspose.Cells | Get list of cell addresses with #DIV/0! result in an Excel file with Aspose.Cells .NET
// Tags: Aspose.Cells detect division by zero error | C# collect formula error cells | Aspose.Cells iterate used range | Excel #DIV/0! error extraction .NET | list error cell addresses Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// The example loads an Excel file with Aspose.Cells, walks through each worksheet's used range, checks each formula cell for a #DIV/0! error by verifying Cell.Type is IsError and StringValue equals "#DIV/0!", adds matching cells to a list, and prints the worksheet name and cell address for every division‑by‑zero error found.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Collection to store cells that contain #DIV/0! errors
        List<Cell> divZeroCells = new List<Cell>();

        try
        {
            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the used range of the worksheet
                AsposeRange usedRange = sheet.Cells.MaxDisplayRange;
                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                // Loop through each cell in the used range
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Process only cells that contain a formula
                        if (cell.IsFormula)
                        {
                            // Check if the formula result is a #DIV/0! error
                            if (cell.Type == CellValueType.IsError && cell.StringValue == "#DIV/0!")
                            {
                                divZeroCells.Add(cell);
                            }
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
            return;
        }

        // Output the addresses of cells with #DIV/0! errors
        foreach (Cell errCell in divZeroCells)
        {
            Console.WriteLine($"Sheet: {errCell.Worksheet.Name}, Cell: {errCell.Name}");
        }

        // If you need to save any changes, uncomment the line below
        // workbook.Save("output.xlsx");
    }
}
