using System;
using Aspose.Cells;

namespace AsposeCellsEnglishFormulaDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook without parsing formulas on open.
            LoadOptions loadOptions = new LoadOptions
            {
                ParsingFormulaOnOpen = false // Skip automatic formula parsing
            };

            // Replace "input.xlsx" with the path to your source workbook.
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set a formula using English function names.
            cells["A1"].Formula = "=SUM(B1:B5)";

            // Optionally put some sample data for the SUM calculation.
            for (int i = 1; i <= 5; i++)
            {
                cells[$"B{i}"].PutValue(i * 10); // B1=10, B2=20, ..., B5=50
            }

            // Calculate all formulas in the workbook.
            workbook.CalculateFormula();

            // Display the calculated result.
            Console.WriteLine("Result of A1 (SUM of B1:B5): " + cells["A1"].Value);

            // Save the workbook with the English formula applied.
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}