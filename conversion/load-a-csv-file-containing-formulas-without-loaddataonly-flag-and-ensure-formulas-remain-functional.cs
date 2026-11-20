// Title: Load a CSV file containing formulas and keep them functional when converting to XLSX with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load CSV with formulas without LoadDataOnly | preserve Excel formulas when converting CSV to XLSX using Aspose.Cells | recalculate formulas after importing a CSV into an Aspose.Cells workbook | convert CSV that includes formulas to XLSX in .NET while keeping calculations active | load CSV with formulas and save as XLSX using Aspose.Cells without losing formula integrity
// Tags: load CSV with formulas Aspose.Cells | preserve formulas during CSV to XLSX conversion | recalculate workbook formulas after CSV import | convert CSV to XLSX without LoadDataOnly | Aspose.Cells LoadOptions CSV format

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates how to load a CSV file that contains Excel formulas using Aspose.Cells for .NET (with LoadDataOnly disabled), recalculate the formulas, and then save the workbook as an XLSX file, ensuring the formulas remain functional.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.csv";
            string outputPath = "output.xlsx";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the CSV file
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Recalculate formulas (if any)
            workbook.CalculateFormula();

            // Save the workbook to XLSX format
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
