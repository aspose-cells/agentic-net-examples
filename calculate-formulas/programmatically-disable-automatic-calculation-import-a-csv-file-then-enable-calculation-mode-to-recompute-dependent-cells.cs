// Title: Disable Auto‑Calc, Import CSV, Re‑Enable Calculation with Aspose.Cells for .NET
// Description: Creates a workbook, sets formula calculation to Manual, imports a CSV file into the first worksheet, switches back to Automatic mode, forces a full recalculation, and saves the result as XLSX using Aspose.Cells.
// Keywords: Aspose.Cells manual calculation | ImportCSV Aspose.Cells | CalcModeType Manual | CalcModeType Automatic | CalculateFormula | bulk CSV import performance | .NET spreadsheet API
// Common Searches: Aspose.Cells turn off formula calculation | Import CSV without recalculating formulas Aspose.Cells | Enable calculation after CSV import .NET | How to use CalculateFormula with Aspose.Cells | Set calculation mode manual then automatic
// Developer Intent: Temporarily suspend automatic formula evaluation, load CSV data efficiently, then reactivate calculation and recompute all dependent formulas before saving.
// Use Cases: Load large CSV datasets without triggering per‑row recalculation, then compute all formulas once. | Batch‑update worksheets from external sources while keeping calculation manual to improve speed. | Create a workbook, import data, ensure final formula results are accurate, and export to XLSX.
// AI Prompts: Generate C# code that disables automatic calculation, imports a CSV file into an Aspose.Cells worksheet, re‑enables calculation, and runs CalculateFormula. | Explain the performance benefits of setting CalcModeType.Manual during bulk CSV import and how to safely revert to Automatic. | Provide a step‑by‑step tutorial for controlling calculation mode, importing CSV data, and saving a workbook with updated formulas using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Creates a workbook, sets formula calculation to Manual, imports a CSV file into the first worksheet, switches back to Automatic mode, forces a full recalculation, and saves the result as XLSX using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet and its cells
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Disable automatic calculation by setting the mode to Manual
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Import CSV data into the worksheet starting at cell A1 (row 0, column 0)
        // Adjust the file path, delimiter, and conversion options as needed
        string csvFilePath = "data.csv";
        cells.ImportCSV(csvFilePath, ",", true, 0, 0); // rule: ImportCSV

        // Re‑enable automatic calculation (or any desired mode)
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

        // Recalculate all formulas now that the data is loaded
        workbook.CalculateFormula(); // rule: CalculateFormula

        // Save the workbook (lifecycle rule: save)
        workbook.Save("result.xlsx", SaveFormat.Xlsx);
    }
}
