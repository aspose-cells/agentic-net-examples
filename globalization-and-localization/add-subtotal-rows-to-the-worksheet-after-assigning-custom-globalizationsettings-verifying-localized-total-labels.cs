// Title: Add a French‑localized subtotal row with a SUM formula to an existing Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a workbook, sets CultureInfo to fr-FR, inserts a row labeled "Sous‑total" after the last data row, and adds a SUM formula for column B with Aspose.Cells. | Demonstrate how to recalculate all formulas after adding a subtotal row in an Aspose.Cells workbook and save the updated file. | Modify the example to place the subtotal in column D and use the German label "Zwischensumme" while preserving the localization steps.
// Common Searches: how to insert a localized subtotal row in Excel using Aspose.Cells C# | Aspose.Cells set workbook culture to French and add sum total row | C# recalculate formulas after adding a new row with Aspose.Cells | determine last populated row and add subtotal label in Aspose.Cells | change subtotal label language in an Aspose.Cells workbook
// Tags: Aspose.Cells subtotal row insertion | Aspose.Cells workbook culture localization | Aspose.Cells insert SUM formula | Aspose.Cells find MaxDataRow | Aspose.Cells recalculate workbook formulas

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file, sets the workbook's CultureInfo to French (fr-FR), writes the label "Sous‑total" in column A after the last data row, adds a SUM formula in column B covering the data range, forces formula recalculation, and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook;
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                    return;
                }

                // Set culture info (French) for the workbook
                workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

                // Localized label for subtotals (used directly, not via Settings)
                const string subtotalLabel = "Sous‑total";

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Determine the last row that contains data (zero‑based index)
                int lastDataRow = sheet.Cells.MaxDataRow;

                // Define the row where the subtotal will be inserted (one row after the last data row)
                int subtotalRowIndex = lastDataRow + 1;

                // Write the localized subtotal label in column A of the new row
                sheet.Cells[subtotalRowIndex, 0].PutValue(subtotalLabel);

                // Insert a SUM formula in column B of the subtotal row
                // Assuming numeric values start at row 2 (Excel row 2, zero‑based index 1) in column B (index 1)
                string sumFormula = $"=SUM(B2:B{lastDataRow + 1})";
                sheet.Cells[subtotalRowIndex, 1].Formula = sumFormula;

                // Recalculate all formulas so the subtotal value is evaluated
                workbook.CalculateFormula();

                // Save the modified workbook
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to: {outputPath}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
