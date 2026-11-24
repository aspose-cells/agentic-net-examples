// Title: Detect and list cells with #DIV/0! errors using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells find cells that return #DIV/0! after calculating formulas | list addresses of division by zero errors in an Excel workbook with Aspose.Cells | enumerate error cells across all worksheets using Aspose.Cells calculate formula | retrieve A1‑style addresses of #DIV/0! cells in a .NET Excel file with Aspose.Cells | detect formula #DIV/0! errors in every sheet using Aspose.Cells C# example
// Tags: Aspose.Cells division by zero error detection | C# extract error cell addresses from Excel | calculate formulas and capture errors Aspose.Cells | iterate worksheet used range Aspose.Cells | output A1 style error addresses Aspose.Cells | save workbook after error analysis Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an Excel workbook (or creates a new one), calculates all formulas, iterates through each worksheet's used range, checks each cell for the #DIV/0! error by examining its type and string value, collects the A1‑style addresses (including sheet name) of those error cells, prints the list, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Load an existing workbook if the file is present; otherwise create a new one.
            Workbook workbook;
            string inputPath = "input.xlsx";

            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load '{inputPath}': {ex.Message}");
                    workbook = new Workbook(); // fallback to empty workbook
                }
            }
            else
            {
                workbook = new Workbook(); // creates an empty workbook
            }

            // Calculate all formulas in the workbook.
            workbook.CalculateFormula();

            // Collect addresses of cells that contain #DIV/0! errors.
            List<string> divZeroCells = new List<string>();

            // Iterate through each worksheet.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the used range of the sheet.
                var usedRange = sheet.Cells.MaxDisplayRange;
                int firstRow = usedRange.FirstRow;
                int firstColumn = usedRange.FirstColumn;
                int rowCount = usedRange.RowCount;
                int columnCount = usedRange.ColumnCount;

                for (int row = firstRow; row < firstRow + rowCount; row++)
                {
                    for (int col = firstColumn; col < firstColumn + columnCount; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Detect #DIV/0! errors by checking the cell type and string representation.
                        if (cell.Type == CellValueType.IsError && cell.StringValue == "#DIV/0!")
                        {
                            // Store the cell address in A1 style with sheet name.
                            divZeroCells.Add($"{sheet.Name}!{cell.Name}");
                        }
                    }
                }
            }

            // Output the collected cell addresses.
            Console.WriteLine("Cells with #DIV/0! error:");
            foreach (string address in divZeroCells)
            {
                Console.WriteLine(address);
            }

            // Save the workbook if needed.
            string outputPath = "output.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
