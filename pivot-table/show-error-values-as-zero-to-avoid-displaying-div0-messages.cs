// Title: Replace #DIV/0! and other Excel error values with zero in an Aspose.Cells workbook using C#
// AI Prompts: Write C# code that calculates all formulas in an Aspose.Cells workbook and then changes every cell of type IsError to 0 before saving the file. | Show how to iterate over the used range of a worksheet with Aspose.Cells, detect error values such as #DIV/0! and replace them with zero.
// Common Searches: Aspose.Cells C# replace #DIV/0! error with 0 | how to set error cells to zero after formula calculation using Aspose.Cells | clean Excel error values in .NET workbook Aspose.Cells | iterate over used cells and fix errors Aspose.Cells C#
// Tags: replace error values Aspose.Cells | Aspose.Cells set error cells to zero | C# Aspose.Cells error handling | Aspose.Cells calculate formulas then clean errors | Excel #DIV/0! replacement Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, inserts values that cause a division‑by‑zero error, calculates formulas to generate the #DIV/0! value, then scans all populated cells, replaces any cell whose type is IsError with 0, and saves the cleaned workbook as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data: A1 = 10, B1 = 0, C1 = =A1/B1 (will cause #DIV/0!)
            cells["A1"].PutValue(10);
            cells["B1"].PutValue(0);
            cells["C1"].Formula = "A1/B1";

            // Calculate formulas so that the error value is generated
            workbook.CalculateFormula();

            // Replace all error values (e.g., #DIV/0!) with zero
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;
            for (int i = 0; i <= maxRow; i++)
            {
                for (int j = 0; j <= maxCol; j++)
                {
                    Cell cell = cells[i, j];
                    if (cell.Type == CellValueType.IsError)
                    {
                        cell.PutValue(0);
                    }
                }
            }

            // Save the workbook
            workbook.Save("Result.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
