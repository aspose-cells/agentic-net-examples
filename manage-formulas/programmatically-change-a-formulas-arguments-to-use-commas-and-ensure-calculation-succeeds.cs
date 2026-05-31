using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFormulaDelimiterDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate some sample values
                cells["A1"].PutValue(10);
                cells["A2"].PutValue(20);
                cells["A3"].PutValue(30);

                // Set a formula that uses semicolon as argument separator (locale dependent)
                // Example: =SUM(A1;A2;A3)
                Cell targetCell = cells["B1"];
                string originalFormula = "=SUM(A1;A2;A3)";

                // Replace semicolons with commas for the current locale
                string correctedFormula = originalFormula.Replace(';', ',');

                // Parse the corrected formula and set it to the cell
                FormulaParseOptions parseOptions = new FormulaParseOptions { Parse = true };
                targetCell.SetFormula(correctedFormula, parseOptions, null);

                // Calculate all formulas in the workbook
                workbook.CalculateFormula();

                // Output the result to verify calculation succeeded
                Console.WriteLine($"Corrected formula: {targetCell.Formula}");
                Console.WriteLine($"Calculated value in B1: {targetCell.Value}");

                // Save the workbook (optional)
                string outputPath = "FormulaDelimiterCorrected.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}