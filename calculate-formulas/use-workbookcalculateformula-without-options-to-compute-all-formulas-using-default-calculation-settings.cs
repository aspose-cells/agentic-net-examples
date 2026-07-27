using System;
using Aspose.Cells;

namespace AsposeCellsCalculateFormulaDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (default constructor)
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells collection
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some data
            cells["A1"].PutValue(5);                 // Simple numeric value
            cells["B1"].Formula = "=A1*2";           // Formula referencing A1
            cells["C1"].Formula = "=B1+10";          // Formula referencing B1

            // Calculate all formulas using default calculation settings
            workbook.CalculateFormula();

            // Output the calculated results to the console
            Console.WriteLine("A1 value: " + cells["A1"].IntValue);
            Console.WriteLine("B1 formula result: " + cells["B1"].IntValue);
            Console.WriteLine("C1 formula result: " + cells["C1"].IntValue);

            // Save the workbook to verify that calculated values are stored
            workbook.Save("CalculatedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}