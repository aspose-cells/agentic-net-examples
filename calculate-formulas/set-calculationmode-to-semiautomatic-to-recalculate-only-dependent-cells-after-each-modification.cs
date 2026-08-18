// Title: Aspose.Cells C# – Set Workbook to Semi‑Automatic Calculation (recalculate dependent cells only)
// Description: Shows how to configure Aspose.Cells for .NET to use CalculationMode AutomaticExceptTable, which mimics a semi‑automatic mode by updating only formulas that depend on changed cells, then saves the workbook.
// Keywords: Aspose.Cells | C# | CalculationMode | AutomaticExceptTable | semi‑automatic calculation | recalculate dependent cells | FormulaSettings | partial workbook recalculation | Excel performance optimization
// Common Searches: Aspose.Cells set calculation mode C# | semi‑automatic calculation mode Aspose.Cells .NET | AutomaticExceptTable example code | recalculate only dependent cells Aspose.Cells | enable partial recalculation in Aspose.Cells workbook
// Developer Intent: Configure a workbook to recalculate only cells that depend on edited values.
// Use Cases: Speed up processing of large spreadsheets by avoiding full‑workbook recalculation. | Create interactive reports that instantly reflect changes in source data without unnecessary calculations. | Build a custom Excel‑like editor that mirrors Excel's semi‑automatic calculation behavior.
// AI Prompts: Explain how Aspose.Cells CalculationMode AutomaticExceptTable works and its performance impact. | Provide a C# snippet that sets CalculationMode to AutomaticExceptTable and demonstrates dependent formula updates. | Compare Automatic, AutomaticExceptTable, and Manual calculation modes in Aspose.Cells, with guidance on when to use each.

using Aspose.Cells;

// Shows how to configure Aspose.Cells for .NET to use CalculationMode AutomaticExceptTable, which mimics a semi‑automatic mode by updating only formulas that depend on changed cells, then saves the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set calculation mode to the closest option for "SemiAutomatic"
        // (recalculates dependent cells after each change)
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.AutomaticExceptTable;

        // Sample data and formula
        sheet.Cells["A1"].PutValue(10);               // Base value
        sheet.Cells["A2"].Formula = "=A1*2";          // Dependent formula

        // Recalculate formulas so dependent cells are updated
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("SemiAutomaticMode.xlsx");
    }
}
