using System;
using Aspose.Cells;

namespace RecalculateFormulasDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // creates a blank workbook

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set initial values and formulas
            cells["A1"].PutValue(5);                 // plain value
            cells["B1"].Formula = "=A1*2";           // formula dependent on A1
            cells["C1"].Formula = "=B1+10";          // formula dependent on B1

            // Modify data that affects formulas
            cells["A1"].PutValue(8); // change A1 from 5 to 8

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Display the updated results
            Console.WriteLine("A1 value: " + cells["A1"].IntValue); // 8
            Console.WriteLine("B1 formula result: " + cells["B1"].IntValue); // 16
            Console.WriteLine("C1 formula result: " + cells["C1"].IntValue); // 26

            // Save the workbook if needed
            workbook.Save("RecalculatedWorkbook.xlsx");
        }
    }
}