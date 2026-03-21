using System;
using Aspose.Cells;

namespace AsposeCellsIFLoadDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file that contains IF formulas.
            string inputPath = "input.xlsx";

            // -----------------------------------------------------------------
            // 1. Create LoadOptions and enable formula parsing on opening.
            //    This ensures that formulas (including IF) are parsed when the
            //    workbook is loaded.
            // -----------------------------------------------------------------
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.ParsingFormulaOnOpen = true; // parse formulas during load

            // -----------------------------------------------------------------
            // 2. Load the workbook with the specified options.
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // -----------------------------------------------------------------
            // 3. Some formulas may still be stored without being parsed
            //    (e.g., if they were saved with Parse = false). Parse all
            //    remaining formulas now.
            // -----------------------------------------------------------------
            workbook.ParseFormulas(false); // do not ignore errors

            // -----------------------------------------------------------------
            // 4. Calculate all formulas so that IF expressions are evaluated.
            // -----------------------------------------------------------------
            workbook.CalculateFormula();

            // -----------------------------------------------------------------
            // 5. Retrieve and display the result of a cell that uses IF.
            //    Adjust the address as needed for your workbook.
            // -----------------------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];
            Cell resultCell = sheet.Cells["B1"]; // assume B1 contains =IF(A1>5,"High","Low")
            Console.WriteLine("Cell B1 formula : " + resultCell.Formula);
            Console.WriteLine("Cell B1 evaluated value : " + resultCell.StringValue);

            // -----------------------------------------------------------------
            // 6. Save the workbook (optional) to verify that the evaluated values
            //    are persisted if desired.
            // -----------------------------------------------------------------
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}