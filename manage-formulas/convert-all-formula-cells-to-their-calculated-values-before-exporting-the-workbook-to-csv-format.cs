// Title: Convert all formula cells to static values before exporting a workbook to CSV with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# calculate all formulas and save as CSV without formulas | How to replace Excel formula cells with their evaluated values before CSV export in .NET | Export XLSX to CSV using Aspose.Cells while preserving only calculated values | Remove formulas from workbook programmatically with Aspose.Cells before saving as CSV | Convert formula cells to values in Aspose.Cells then generate a CSV file
// Tags: calculate formulas Aspose.Cells | replace formula with value Aspose.Cells | export worksheet to CSV Aspose.Cells | static value CSV conversion .NET | remove Excel formulas programmatically

using System;
using System.IO;
using Aspose.Cells;

// The example loads an XLSX workbook, forces calculation of all formulas, replaces each formula cell with its evaluated result, sets the first worksheet as active, and saves the sheet as a CSV file using Aspose.Cells for .NET, while handling missing files and password‑protected workbooks.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.csv";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook; if the file is password‑protected, an exception will be caught below
            Workbook workbook = new Workbook(inputPath);

            // Ensure all formulas are calculated
            workbook.CalculateFormula();

            // Replace each formula with its calculated value
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                foreach (Cell cell in cells)
                {
                    if (cell.IsFormula)
                    {
                        object calculatedValue = cell.Value;
                        cell.Formula = string.Empty;          // Remove the formula
                        cell.PutValue(calculatedValue);       // Set the calculated value
                    }
                }
            }

            // Set the first worksheet as the active sheet for CSV export
            workbook.Worksheets.ActiveSheetIndex = 0;

            // Export the active sheet to CSV
            workbook.Save(outputPath, SaveFormat.Csv);

            Console.WriteLine($"Workbook has been successfully saved as \"{outputPath}\".");
        }
        catch (CellsException ex) when (ex.Message.Contains("password"))
        {
            Console.WriteLine("Error: The workbook is password protected. Provide a valid password to load it.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
