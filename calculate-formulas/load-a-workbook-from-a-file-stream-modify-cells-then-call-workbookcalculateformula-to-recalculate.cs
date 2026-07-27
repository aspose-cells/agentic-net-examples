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

            // Open the file as a FileStream for reading
            using (FileStream fileStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
            {
                // Load the workbook from the stream using the provided constructor
                Workbook workbook = new Workbook(fileStream);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Modify some cells (example: set values and a formula)
                cells["A1"].PutValue(10);               // Set a numeric value
                cells["B1"].PutValue(20);               // Set another numeric value
                cells["C1"].Formula = "=A1+B1";         // Set a formula that depends on A1 and B1

                // Recalculate all formulas in the workbook
                workbook.CalculateFormula();

                // Optionally, save the updated workbook back to a file
                workbook.Save("output.xlsx", SaveFormat.Xlsx);
            }

            Console.WriteLine("Workbook loaded, modified, and formulas recalculated successfully.");
        }
    }
}