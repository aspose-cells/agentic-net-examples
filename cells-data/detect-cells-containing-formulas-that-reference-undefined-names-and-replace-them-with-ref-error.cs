// Title: C# Aspose.Cells: Detect undefined‑name formulas and replace them with #REF! error
// Description: Loads an Excel workbook, runs CalculateFormula to expose #NAME? errors, scans each worksheet’s used range, finds formula cells that evaluate to an error, and substitutes the formula with a "#REF!" placeholder before saving the file.
// Keywords: Aspose.Cells C# replace formula error | detect undefined name Excel | convert #NAME? to #REF! | iterate cells Aspose.Cells | C# Excel error handling | batch fix formula errors
// Common Searches: how to replace #NAME? with #REF! using Aspose.Cells | C# scan workbook for error formulas Aspose.Cells | Aspose.Cells change formula errors to #REF! | detect undefined names in Excel C#
// Developer Intent: Identify formula cells that reference undefined names and replace them with a #REF! error value.
// Use Cases: Sanitize workbooks before distribution by converting all undefined‑name errors to a uniform #REF! marker. | Prepare Excel files for data import pipelines that cannot handle #NAME? errors. | Automate batch processing of multiple spreadsheets to ensure no missing‑name references remain.
// AI Prompts: Generate C# code that opens a workbook with Aspose.Cells, runs CalculateFormula, finds cells where IsFormula is true and Type equals IsError, and sets the cell value to "#REF!". | Provide a C# snippet that logs the addresses (e.g., A1, B2) of every formula cell changed to "#REF!" while processing a workbook with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, runs CalculateFormula to expose #NAME? errors, scans each worksheet’s used range, finds formula cells that evaluate to an error, and substitutes the formula with a "#REF!" placeholder before saving the file.
class Program
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

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Calculate all formulas so that errors (e.g., #NAME?) are evaluated
            workbook.CalculateFormula();

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range of the sheet
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Scan every cell in the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Process only formula cells that resulted in an error after calculation
                        if (cell.IsFormula && cell.Type == CellValueType.IsError)
                        {
                            // Replace the erroneous formula with a #REF! error representation
                            // Using a string value to avoid reliance on PutErrorValue API variations
                            cell.PutValue("#REF!");
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
