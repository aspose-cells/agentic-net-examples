// Title: Import CSV with embedded Excel formulas using Aspose.Cells for .NET and keep them functional
// AI Prompts: Configure TxtLoadOptions.HasFormula = true so that cells beginning with '=' in a CSV are treated as Excel formulas during import. | After importing the CSV data, invoke Workbook.CalculateFormula() to evaluate all formulas before saving the workbook. | Set TxtLoadOptions.Separator and ConvertNumericData as required while preserving formulas when converting CSV to XLSX.
// Common Searches: Aspose.Cells C# load CSV with formulas without LoadDataOnly flag | How to keep '=SUM' formulas active when importing CSV into Aspose.Cells workbook | Enable formula detection for CSV to XLSX conversion using Aspose.Cells .NET | Preserve Excel formulas during CSV import with TxtLoadOptions in Aspose.Cells | Calculate imported CSV formulas before saving as XLSX in Aspose.Cells
// Tags: Csv import with formula detection Aspose.Cells | TxtLoadOptions.HasFormula property | ImportCSV method preserving formulas | Workbook.CalculateFormula after CSV import | CSV to XLSX conversion preserving Excel formulas

using System;
using Aspose.Cells;

namespace AsposeCellsCsvFormulaDemo
{
    // The example demonstrates loading a CSV that contains formula strings, enabling TxtLoadOptions.HasFormula to treat leading '=' as formulas, importing the data into a new workbook, recalculating all formulas, and saving the result as an XLSX file while keeping the formulas functional.
    class Program
    {
        static void Main()
        {
            // Path to the CSV file that contains formulas (e.g., cells with values like "=SUM(A1:A3)")
            string csvPath = "input.csv";

            // Create TxtLoadOptions and enable formula detection.
            TxtLoadOptions txtOptions = new TxtLoadOptions();
            txtOptions.HasFormula = true;          // Treat text starting with '=' as a formula.
            txtOptions.Separator = ',';            // Define the CSV delimiter.
            txtOptions.ConvertNumericData = true;  // Convert numeric strings to numbers.

            // Create an empty workbook.
            Workbook workbook = new Workbook();

            // Import the CSV data into the first worksheet starting at cell A1 (row 0, column 0).
            workbook.Worksheets[0].Cells.ImportCSV(csvPath, txtOptions, 0, 0);

            // Calculate all formulas that were imported.
            workbook.CalculateFormula();

            // Save the workbook to an Excel file; formulas will be functional.
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
