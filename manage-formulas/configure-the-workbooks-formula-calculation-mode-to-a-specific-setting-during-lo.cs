using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create load options (you can adjust other options here if needed)
        LoadOptions loadOptions = new LoadOptions();

        // Load the XLSX workbook with the specified load options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Configure the formula calculation mode for the loaded workbook.
        // Options: Automatic, AutomaticExceptTable, Manual
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.AutomaticExceptTable;

        // Save the workbook (optional, depending on whether you need to persist the change)
        workbook.Save("output.xlsx");
    }
}