using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFormulaRecalcDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the existing Excel file
            string inputPath = "input.xlsx";

            // Load the workbook from a file stream using the provided constructor
            using (FileStream inputStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
            {
                Workbook workbook = new Workbook(inputStream);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Modify some cells
                sheet.Cells["A1"].PutValue(10);               // Set a numeric value
                sheet.Cells["B1"].PutValue(20);               // Set another numeric value
                sheet.Cells["C1"].Formula = "=A1+B1";         // Set a formula that depends on the modified cells

                // Recalculate all formulas in the workbook
                workbook.CalculateFormula();

                // Optionally, save the updated workbook to a new file
                workbook.Save("output.xlsx", SaveFormat.Xlsx);
            }

            Console.WriteLine("Workbook loaded, cells modified, and formulas recalculated successfully.");
        }
    }
}