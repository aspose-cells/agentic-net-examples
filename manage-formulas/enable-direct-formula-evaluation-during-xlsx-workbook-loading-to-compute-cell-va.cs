using System;
using Aspose.Cells;

namespace AsposeCellsFormulaEvaluation
{
    class Program
    {
        static void Main()
        {
            // Specify load options to parse formulas when the workbook is opened.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.ParsingFormulaOnOpen = true; // Enable formula parsing on load

            // Load the XLSX workbook with the defined options.
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Immediately evaluate all formulas after loading.
            workbook.CalculateFormula();

            // Example: read a calculated cell value.
            Worksheet sheet = workbook.Worksheets[0];
            Cell resultCell = sheet.Cells["B1"]; // Adjust the address as needed
            Console.WriteLine($"Calculated value in B1: {resultCell.Value}");
        }
    }
}