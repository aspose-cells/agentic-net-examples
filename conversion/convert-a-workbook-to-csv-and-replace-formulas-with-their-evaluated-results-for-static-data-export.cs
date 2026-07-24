// Title: C# – Export Excel to CSV with Evaluated Formulas Using Aspose.Cells
// Description: Loads an .xlsx workbook, forces a full formula calculation, strips all formulas leaving only their computed values, and saves the first worksheet as a value‑only CSV file.
// Keywords: Aspose.Cells | C# | .NET | Excel to CSV | evaluate formulas | remove formulas | static CSV export | Workbook.CalculateFormula | Cells.RemoveFormulas | SaveFormat.Csv
// Common Searches: Aspose.Cells export Excel to CSV with formula results | C# remove formulas before saving as CSV | calculate all formulas then convert workbook to CSV | static CSV export from Excel using Aspose.Cells | how to strip formulas in Aspose.Cells
// Developer Intent: Create a CSV file from an Excel workbook where every formula is pre‑calculated and replaced by its value.
// Use Cases: Generate a plain‑text report from a calculation‑intensive spreadsheet for downstream analytics. | Convert user‑filled Excel templates to value‑only CSV files for systems that reject formulas. | Automate data extraction from Excel to CSV for integration with legacy import pipelines.
// AI Prompts: Write C# code with Aspose.Cells that loads an .xlsx, evaluates all formulas, removes them, and saves the result as a CSV. | Explain the effect of Cells.RemoveFormulas on the CSV output produced by Aspose.Cells. | Provide a solution to export each worksheet of a workbook to separate CSV files after evaluating all formulas.

using System;
using Aspose.Cells;

// Loads an .xlsx workbook, forces a full formula calculation, strips all formulas leaving only their computed values, and saves the first worksheet as a value‑only CSV file.
class Program
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Calculate all formulas so that their results are up‑to‑date
        workbook.CalculateFormula();

        // Replace formulas with their evaluated values for static export
        workbook.Worksheets[0].Cells.RemoveFormulas();

        // Save the workbook as CSV (values only, no formulas)
        workbook.Save("output.csv", SaveFormat.Csv);

        Console.WriteLine("Workbook has been exported to CSV with formulas evaluated.");
    }
}
