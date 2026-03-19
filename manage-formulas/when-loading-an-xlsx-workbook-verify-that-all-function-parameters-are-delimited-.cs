using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class VerifyFormulaParameterSeparator
    {
        public static void Run()
        {
            // Path to the XLSX workbook to be loaded
            string workbookPath = "input.xlsx";

            // Create load options and ensure formulas are parsed on open
            LoadOptions loadOptions = new LoadOptions
            {
                ParsingFormulaOnOpen = true   // parse formulas while loading
            };

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(workbookPath, loadOptions);

            // Retrieve the list separator used for function parameters in the workbook's globalization settings
            string listSeparatorString = workbook.Settings.CultureInfo.TextInfo.ListSeparator;
            char listSeparator = string.IsNullOrEmpty(listSeparatorString) ? ',' : listSeparatorString[0];

            // Initial verification based on the global list separator
            bool allParametersUseComma = listSeparator == ',';
            Console.WriteLine($"Global list separator is '{listSeparator}'. Expected ',': {allParametersUseComma}");

            // Additional verification: scan each cell that contains a formula and ensure commas are used
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                foreach (Cell cell in cells)
                {
                    if (cell.IsFormula)
                    {
                        // Simple check: if the formula contains a semicolon, it likely uses a non‑comma separator
                        if (cell.Formula.Contains(";"))
                        {
                            allParametersUseComma = false;
                            Console.WriteLine($"Cell {cell.Name} uses ';' as a separator in formula: {cell.Formula}");
                        }
                    }
                }
            }

            // Final result
            if (allParametersUseComma)
            {
                Console.WriteLine("All function parameters in the workbook are delimited by commas.");
            }
            else
            {
                Console.WriteLine("Some function parameters are not delimited by commas.");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            VerifyFormulaParameterSeparator.Run();
        }
    }
}