using System;
using System.IO;
using Aspose.Cells;

class ConvertFormulasToEnglish
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains formulas in a locale‑specific format
            Workbook wb = new Workbook(inputPath);

            // Force the workbook to use US English settings (comma as argument separator)
            wb.Settings.Region = CountryCode.USA;

            // Parse any formulas that were set without parsing (e.g., when loaded from a template)
            wb.ParseFormulas(false);

            // Convert each formula to English (US) syntax
            foreach (Worksheet sheet in wb.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate over all cells that have a formula
                foreach (Cell cell in cells)
                {
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        // Retrieve the formula in the locale representation (may use semicolons, etc.)
                        string localFormula = cell.FormulaLocal;

                        // Re‑assign the formula; Aspose.Cells stores it in English syntax (commas) after parsing
                        cell.Formula = localFormula;
                    }
                }
            }

            // Calculate all formulas to ensure they evaluate correctly after conversion
            wb.CalculateFormula();

            // Validate that no cell contains an error after calculation
            bool hasError = false;
            foreach (Worksheet sheet in wb.Worksheets)
            {
                foreach (Cell cell in sheet.Cells)
                {
                    // Check for error cells using the Cell.Type property
                    if (cell.Type == CellValueType.IsError)
                    {
                        Console.WriteLine($"Error in sheet '{sheet.Name}' cell {cell.Name}: {cell.StringValue}");
                        hasError = true;
                    }
                }
            }

            if (!hasError)
            {
                Console.WriteLine("All formulas converted to English (US) syntax and calculated successfully.");
            }

            // Save the workbook with the converted formulas
            wb.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}