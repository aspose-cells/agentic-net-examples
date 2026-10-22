// Title: Load a CSV file with embedded formulas using Aspose.Cells for .NET and keep the formulas functional when saving to XLSX
// Common Searches: c# load csv file that contains Excel formulas with Aspose.Cells | preserve and recalculate formulas after converting CSV to XLSX using Aspose.Cells | how to use LoadOptions for CSV import without LoadDataOnly in Aspose.Cells | Aspose.Cells calculate formulas after importing a CSV workbook | convert CSV to XLSX while keeping formulas intact in .NET
// Tags: load CSV with formulas Aspose.Cells | CSV to XLSX conversion with active formulas | Workbook.CalculateFormula after CSV import | LoadOptions CSV without LoadDataOnly flag | Aspose.Cells formula recalculation workflow | saving workbook as XLSX with active formulas

using Aspose.Cells;
using System;
using System.IO;

// The example verifies the input CSV file, loads it into an Aspose.Cells Workbook using LoadOptions set to CSV (without the LoadDataOnly flag), triggers formula recalculation with CalculateFormula, and then saves the workbook as an XLSX file, ensuring that any embedded formulas remain functional.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.csv";
            string outputPath = "output.xlsx";

            // Verify that the input CSV file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the CSV file with CSV load format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Recalculate any formulas present
            workbook.CalculateFormula();

            // Save the workbook as XLSX, preserving formulas
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
