// Title: How to control calculation mode while importing a CSV and saving as XLSX using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load CSV without triggering automatic calculation | programmatically enable formula recalculation after importing CSV with Aspose.Cells | disable workbook auto‑calculate, import CSV, then recalculate formulas in .NET | convert CSV to XLSX and force CalculateFormula using Aspose.Cells API | set calculation mode before saving workbook after CSV import in C#
// Tags: disable automatic calculation Aspose.Cells | import CSV workbook Aspose.Cells C# | enable calculation mode Aspose.Cells | calculate formulas after CSV import | convert CSV to XLSX Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for an input CSV file, loads it into an Aspose.Cells Workbook with CSV LoadOptions, optionally controls the calculation mode, forces a full formula recalculation using CalculateFormula, and saves the result as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Verify that the input CSV file exists to avoid FileNotFoundException
            const string inputPath = "input.csv";
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the CSV file into a workbook using LoadOptions for CSV format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Force a full recalculation of formulas (if any)
            workbook.CalculateFormula();

            // Save the result in XLSX format
            const string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
