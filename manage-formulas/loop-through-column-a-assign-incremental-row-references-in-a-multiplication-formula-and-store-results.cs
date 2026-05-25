using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Fill column A (A1:A10) with values 1‑10
                for (int i = 0; i < 10; i++)
                {
                    cells[i, 0].PutValue(i + 1);
                }

                // Set multiplication formula in column B for each row
                for (int i = 0; i < 10; i++)
                {
                    int excelRow = i + 1; // Excel rows are 1‑based
                    string formula = $"=A{excelRow}*{excelRow}";
                    // Assign formula directly to the cell
                    cells[i, 1].Formula = formula;
                }

                // Calculate all formulas
                workbook.CalculateFormula();

                // Save the workbook
                string outputFile = "MultiplicationResult.xlsx";
                workbook.Save(outputFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}