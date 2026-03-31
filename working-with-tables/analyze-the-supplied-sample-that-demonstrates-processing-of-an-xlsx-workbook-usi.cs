using System;
using Aspose.Cells;

namespace AsposeCellsProcessingDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLSX file
            string sourcePath = "InputWorkbook.xlsx";

            // Load the workbook using the constructor that accepts a file path
            Workbook workbook = new Workbook(sourcePath);

            // Access the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Read existing values from cells A1 and B1
            object valueA1 = sheet.Cells["A1"].Value;
            object valueB1 = sheet.Cells["B1"].Value;
            Console.WriteLine($"Original A1: {valueA1}, B1: {valueB1}");

            // Write new data to the worksheet
            sheet.Cells["C1"].PutValue("Processed");
            sheet.Cells["D1"].PutValue(DateTime.Now);

            // Add a simple formula in cell E1 that sums B1 and D1 (if numeric)
            sheet.Cells["E1"].Formula = "=SUM(B1,D1)";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Retrieve the result of the formula
            object formulaResult = sheet.Cells["E1"].Value;
            Console.WriteLine($"Formula result in E1: {formulaResult}");

            // Save the modified workbook to a new file
            string outputPath = "ProcessedWorkbook.xlsx";
            workbook.Save(outputPath);

            // Clean up
            workbook.Dispose();

            Console.WriteLine($"Workbook processing complete. Saved to '{outputPath}'.");
        }
    }
}