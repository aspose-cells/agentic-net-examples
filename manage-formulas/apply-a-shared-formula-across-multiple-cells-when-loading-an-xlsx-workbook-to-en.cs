using System;
using Aspose.Cells;

namespace SharedFormulaDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Example: ensure column A has data for the formula to reference
            // (If the workbook already contains data, this loop can be omitted)
            for (int i = 0; i < 10; i++)
            {
                cells[i, 0].PutValue(i + 1); // A1:A10 = 1..10
            }

            // Apply a shared formula to column B (B1:B10) that multiplies column A by 2
            // Using the overload with FormulaParseOptions for future‑proofing
            FormulaParseOptions parseOptions = new FormulaParseOptions(); // default options
            cells["B1"].SetSharedFormula("=A1*2", 10, 1, parseOptions);

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Optional: display a few results to verify
            Console.WriteLine("Results after applying shared formula:");
            for (int i = 0; i < 5; i++)
            {
                Console.WriteLine($"A{i + 1} = {cells[i, 0].Value}, B{i + 1} = {cells[i, 1].Value}");
            }

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}