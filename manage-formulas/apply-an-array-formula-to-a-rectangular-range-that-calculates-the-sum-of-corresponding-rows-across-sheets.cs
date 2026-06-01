using System;
using Aspose.Cells;

namespace AsposeCellsArrayFormulaDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Prepare source worksheets with sample data
            // -------------------------------------------------
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";

            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3"); // Destination for array formula

            // Fill Sheet1 and Sheet2 with 5 rows of numeric data in column A
            for (int i = 0; i < 5; i++)
            {
                sheet1.Cells[i, 0].PutValue(i + 1);          // Sheet1!A1:A5 = 1,2,3,4,5
                sheet2.Cells[i, 0].PutValue((i + 1) * 10);   // Sheet2!A1:A5 = 10,20,30,40,50
            }

            // -------------------------------------------------
            // Apply an array formula on Sheet3 that sums the
            // corresponding rows from Sheet1 and Sheet2
            // -------------------------------------------------
            // Formula: =Sheet1!A1:A5 + Sheet2!A1:A5
            // It will spill the result into 5 rows (rowNumber = 5) and 1 column
            string arrayFormula = "=Sheet1!A1:A5 + Sheet2!A1:A5";

            // Set the array formula starting at cell A1 of Sheet3
            sheet3.Cells["A1"].SetArrayFormula(arrayFormula, 5, 1);

            // -------------------------------------------------
            // Calculate all formulas so that the array results are populated
            // -------------------------------------------------
            workbook.CalculateFormula();

            // Optional: display the calculated results in the console
            Console.WriteLine("Array formula results in Sheet3 (A1:A5):");
            for (int i = 0; i < 5; i++)
            {
                Console.WriteLine($"Row {i + 1}: {sheet3.Cells[i, 0].Value}");
            }

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            workbook.Save("ArrayFormulaDemo.xlsx");
        }
    }
}