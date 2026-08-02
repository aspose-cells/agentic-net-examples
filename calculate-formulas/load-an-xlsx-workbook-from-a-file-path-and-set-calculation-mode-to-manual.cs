using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the existing XLSX file
        string filePath = "input.xlsx";

        // Load the workbook from the specified file path
        Workbook workbook = new Workbook(filePath);

        // Set the calculation mode to Manual
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Save the workbook if you want to persist the setting
        workbook.Save("output_manual.xlsx", SaveFormat.Xlsx);
    }
}