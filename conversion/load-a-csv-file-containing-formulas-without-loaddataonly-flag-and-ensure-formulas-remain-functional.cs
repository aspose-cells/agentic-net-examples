// Title: Load CSV with Formulas in Aspose.Cells for .NET – Preserve and Calculate Without LoadDataOnly
// Description: Demonstrates how to load a CSV file that contains formula strings using Aspose.Cells for .NET, keep the formulas active by omitting the LoadDataOnly flag, parse and calculate them, and save the result as an XLSX workbook.
// Keywords: Aspose.Cells | CSV formulas | LoadOptions | LoadDataOnly false | ParseFormulas | CalculateFormula | .NET | C# | Excel conversion | preserve formulas
// Common Searches: Aspose.Cells load CSV with formulas | keep formulas when converting CSV to Excel .NET | LoadOptions CSV without LoadDataOnly | ParseFormulas after CSV load Aspose | calculate formulas from CSV using Aspose.Cells
// Developer Intent: Load a CSV file that contains formula expressions, ensure the formulas remain functional, evaluate them, and export the workbook to Excel format.
// Use Cases: Convert a CSV export that includes embedded calculations into a fully functional Excel workbook. | Automate processing of financial CSV feeds where formulas must be evaluated before reporting. | Generate Excel reports from CSV sources that contain formula strings, preserving calculation logic.
// AI Prompts: Write C# code with Aspose.Cells to load a CSV containing formulas, parse and calculate them, then save as XLSX. | Explain the effect of the LoadDataOnly flag on formula handling when loading CSV files with Aspose.Cells. | Provide robust error handling for invalid formulas during Workbook.ParseFormulas after a CSV load.

using System;
using Aspose.Cells;

// Demonstrates how to load a CSV file that contains formula strings using Aspose.Cells for .NET, keep the formulas active by omitting the LoadDataOnly flag, parse and calculate them, and save the result as an XLSX workbook.
class LoadCsvWithFormulas
{
    static void Main()
    {
        // Path to the CSV file that contains formulas (e.g., cells with values like "=SUM(A1:B1)")
        string csvPath = "input.csv";

        // Create load options specifying CSV format.
        // No LoadDataOnly flag is set (default is false), so formulas will be loaded.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);

        // Load the CSV file into a workbook using the specified options.
        Workbook workbook = new Workbook(csvPath, loadOptions);

        // Parse any formulas that were not parsed during the load operation.
        // The 'false' argument means do not ignore errors; an exception will be thrown for invalid formulas.
        workbook.ParseFormulas(false);

        // Calculate the formulas so that cell values reflect the computed results.
        workbook.CalculateFormula();

        // Save the workbook to an Excel file to verify that formulas remain functional.
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
