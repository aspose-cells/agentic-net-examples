using System;
using Aspose.Cells;

namespace SharedFormulaDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string inputPath = "input.xlsx";

            // Create load options and disable formula parsing on open.
            // This allows us to set shared formulas before any automatic parsing occurs.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.ParsingFormulaOnOpen = false;

            // Load the workbook with the specified options.
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Define a shared formula that multiplies the value in column A by 2.
            // Apply it to a range starting at B1, covering 10 rows and 1 column.
            sheet.Cells["B1"].SetSharedFormula("=A1*2", 10, 1);

            // Optionally calculate all formulas so that the results are stored.
            workbook.CalculateFormula();

            // Save the modified workbook.
            workbook.Save("output.xlsx");
        }
    }
}