using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SetSharedFormulaOnLoadDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Configure load options to skip formula parsing on open.
            LoadOptions loadOptions = new LoadOptions
            {
                ParsingFormulaOnOpen = false
            };

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Example data in column A (required for the shared formula)
            for (int i = 0; i < 5; i++)
            {
                cells[i, 0].PutValue(i + 1); // A1-A5 = 1..5
            }

            // Set a shared formula in column B starting from B1 for 5 rows.
            // The formula multiplies the corresponding A column value by 2.
            cells["B1"].SetSharedFormula("=A1*2", 5, 1);

            // After setting the shared formula, calculate all formulas.
            workbook.CalculateFormula();

            // Verify the results (optional)
            for (int i = 0; i < 5; i++)
            {
                Console.WriteLine($"B{i + 1} value: {cells[i, 1].Value}");
            }

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}