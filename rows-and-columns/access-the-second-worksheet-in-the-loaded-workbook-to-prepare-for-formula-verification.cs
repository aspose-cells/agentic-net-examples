// Title: C# – Access Second Worksheet, Verify Formula, Recalculate & Save with Aspose.Cells
// Description: Load an existing Excel file, confirm it has at least two sheets, select the second worksheet (index 1), optionally set it as the active sheet, recalculate all formulas using CalculationOptions, read the formula and computed value of cell B2, and save the updated workbook. Includes file‑existence checks and worksheet‑count validation for robust error handling.
// Keywords: Aspose.Cells second worksheet | C# read cell formula | Aspose.Cells calculate formulas | verify Excel formula .NET | set active sheet Aspose.Cells | Workbook.CalculateFormula | load workbook Aspose.Cells | save workbook Aspose.Cells | Excel file validation C#
// Common Searches: Aspose.Cells get second worksheet C# | read and verify cell formula Aspose.Cells | recalculate all formulas in workbook Aspose | set active sheet index Aspose.Cells .NET | save modified Excel file with Aspose.Cells
// Developer Intent: Load a workbook, target its second sheet, recalculate formulas, inspect a specific cell’s formula and value, then write the changes to a new file.
// Use Cases: Validate that an Excel file contains the expected number of worksheets before processing. | Programmatically activate the second worksheet and ensure all formulas are up‑to‑date. | Extract and display the formula and evaluated result of a key cell (e.g., B2) for verification. | Save the workbook after verification to a separate output file.
// AI Prompts: Generate C# code using Aspose.Cells to open a workbook, select the second worksheet, recalculate all formulas, and print the formula and value of cell B2. | Provide a robust error‑handling template for loading an Excel file, checking worksheet count, and accessing a worksheet by index with Aspose.Cells. | Create a unit test in .NET that confirms Workbook.CalculateFormula updates the value of a formula cell on the second worksheet.

using System;
using System.IO;
using Aspose.Cells;

// Load an existing Excel file, confirm it has at least two sheets, select the second worksheet (index 1), optionally set it as the active sheet, recalculate all formulas using CalculationOptions, read the formula and computed value of cell B2, and save the updated workbook. Includes file‑existence checks and worksheet‑count validation for robust error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "verified_output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Verify that the workbook contains at least two worksheets
            if (workbook.Worksheets.Count < 2)
            {
                Console.WriteLine("The workbook does not contain a second worksheet.");
                return;
            }

            // Access the second worksheet (zero‑based index)
            Worksheet secondWorksheet = workbook.Worksheets[1];

            // Optionally set the second worksheet as the active sheet
            workbook.Worksheets.ActiveSheetIndex = secondWorksheet.Index;

            // Calculate all formulas in the workbook using CalculationOptions
            CalculationOptions calcOptions = new CalculationOptions();
            workbook.CalculateFormula(calcOptions);

            // Example verification: read a specific cell's formula and its calculated value
            Cell cellToVerify = secondWorksheet.Cells["B2"];
            Console.WriteLine($"Cell B2 formula: {cellToVerify.Formula}");
            Console.WriteLine($"Calculated value: {cellToVerify.Value}");

            // Save the workbook after verification
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
