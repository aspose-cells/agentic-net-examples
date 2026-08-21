// Title: Aspose.Cells .NET – Set Workbook Calculation Mode to Automatic After Data Import
// Description: C# example that creates a workbook, simulates data import into each worksheet, then sets the workbook's FormulaSettings.CalculationMode to Automatic (the setting applies globally) and saves the file.
// Keywords: Aspose.Cells calculation mode automatic | C# set workbook formula calculation | Aspose.Cells iterate worksheets | FormulaSettings.CalculationMode .NET | automatic formula recalculation after import
// Common Searches: how to enable automatic calculation in Aspose.Cells .NET | set calculation mode to automatic for all worksheets Aspose.Cells | Aspose.Cells change formula settings after data load | C# Aspose.Cells automatic recalculation example
// Developer Intent: Enable automatic formula recalculation for a workbook after programmatically importing data into its worksheets.
// Use Cases: Import large data sets into multiple sheets and ensure formulas update instantly. | Prepare a workbook for end‑user editing by turning on automatic calculation after bulk data population. | Reset calculation mode to Automatic in batch‑processed workbooks before distribution.
// AI Prompts: Write C# code using Aspose.Cells that imports data into every worksheet and then sets the workbook's calculation mode to Automatic. | Show how to configure FormulaSettings.CalculationMode to Automatic after filling cells in a multi‑sheet workbook. | Explain the difference between Manual and Automatic calculation modes in Aspose.Cells and how to switch them after data manipulation.

using Aspose.Cells;

// C# example that creates a workbook, simulates data import into each worksheet, then sets the workbook's FormulaSettings.CalculationMode to Automatic (the setting applies globally) and saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook (empty with one default worksheet)
        Workbook workbook = new Workbook();

        // -------------------- Data import simulation --------------------
        // Here you would import your data into each worksheet.
        // For demonstration, we fill each worksheet with simple values.
        foreach (Worksheet ws in workbook.Worksheets)
        {
            ws.Cells["A1"].PutValue("Imported");
            ws.Cells["B1"].PutValue(100);
        }
        // ----------------------------------------------------------------

        // After completing data imports, set the calculation mode to Automatic.
        // The CalculationMode property is defined at the workbook level,
        // but we iterate over worksheets as requested.
        foreach (Worksheet ws in workbook.Worksheets)
        {
            // Setting the mode (the same setting applies to the whole workbook)
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;
        }

        // Save the workbook to a file.
        workbook.Save("Result.xlsx", SaveFormat.Xlsx);
    }
}
