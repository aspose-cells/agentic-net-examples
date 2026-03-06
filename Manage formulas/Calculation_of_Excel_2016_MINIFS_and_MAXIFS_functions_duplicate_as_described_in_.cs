using System;
using Aspose.Cells;

namespace AsposeCellsMinMaxIfsDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the existing XLSX file (replace with actual path)
            string inputPath = "InputData.xlsx";

            // Load the workbook (lifecycle: load)
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // Prepare sample data for MINIFS and MAXIFS demonstration
            // Column A: Values to evaluate
            // Column B: Criteria range (e.g., categories)
            // ------------------------------------------------------------
            cells["A1"].PutValue("Value");
            cells["B1"].PutValue("Category");
            cells["A2"].PutValue(10);
            cells["A3"].PutValue(20);
            cells["A4"].PutValue(30);
            cells["A5"].PutValue(40);
            cells["B2"].PutValue("X");
            cells["B3"].PutValue("Y");
            cells["B4"].PutValue("X");
            cells["B5"].PutValue("Y");

            // Insert MINIFS formula: minimum value where Category = "X"
            cells["C2"].Formula = "=MINIFS(A2:A5, B2:B5, \"X\")";

            // Insert MAXIFS formula: maximum value where Category = "Y"
            cells["C3"].Formula = "=MAXIFS(A2:A5, B2:B5, \"Y\")";

            // ------------------------------------------------------------
            // Calculate all formulas in the workbook (lifecycle: calculate)
            // ------------------------------------------------------------
            workbook.CalculateFormula();

            // Retrieve and display the results
            Console.WriteLine("MINIFS result (Category = X): " + cells["C2"].Value);
            Console.WriteLine("MAXIFS result (Category = Y): " + cells["C3"].Value);

            // ------------------------------------------------------------
            // Save the workbook with calculated results (lifecycle: save)
            // ------------------------------------------------------------
            string outputPath = "OutputData.xlsx";
            workbook.Save(outputPath);
        }
    }
}