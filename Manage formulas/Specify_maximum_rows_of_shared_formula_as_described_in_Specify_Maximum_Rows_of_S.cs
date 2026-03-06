using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main()
        {
            MaxRowsOfSharedFormulaDemo.Run();
        }
    }

    public class MaxRowsOfSharedFormulaDemo
    {
        public static void Run()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Specify the maximum number of rows that a shared formula can span
            workbook.Settings.MaxRowsOfSharedFormula = 500;

            // Access the first worksheet and its cells collection
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data in column A (required for the formula)
            for (int i = 0; i < 600; i++)
            {
                cells[i, 0].PutValue(i + 1); // A1, A2, ..., A600
            }

            // Set a shared formula starting at B1 that attempts to fill 600 rows
            cells["B1"].SetSharedFormula("=A1", 600, 1);

            // Verify the formula in the last row that should contain it (row 500)
            Console.WriteLine("Formula in B500: " + cells[499, 1].Formula);

            // Verify that rows beyond the limit do not have the formula
            Console.WriteLine("Formula in B501: " + (cells[500, 1].Formula ?? "null"));

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}