using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class SpecifyMaxRowsOfSharedFormulaDemo
    {
        static void Main()
        {
            // Load an existing XLSX workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            LoadOptions loadOptions = new LoadOptions
            {
                // Skip formula parsing on load for faster loading (optional)
                ParsingFormulaOnOpen = false
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Set the maximum number of rows that a shared formula can span.
            // This must be set before creating a shared formula that exceeds the default limit.
            workbook.Settings.MaxRowsOfSharedFormula = 2000; // example limit

            // Access the first worksheet and its cells collection
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data in column A (required for the formula)
            for (int i = 0; i < 1500; i++)
            {
                cells[i, 0].PutValue(i + 1); // A1..A1500 = 1..1500
            }

            // Set a shared formula in column B that spans 1500 rows.
            // The formula references the corresponding cell in column A.
            cells["B1"].SetSharedFormula("=A1", 1500, 1);

            // Verify that the formula was applied to the last cell in the range
            Console.WriteLine("Formula in B1500: " + cells["B1500"].Formula);

            // Calculate formulas to populate values
            workbook.CalculateFormula();

            // Output a few sample results
            Console.WriteLine("B1 value: " + cells["B1"].Value);
            Console.WriteLine("B1500 value: " + cells["B1500"].Value);

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}