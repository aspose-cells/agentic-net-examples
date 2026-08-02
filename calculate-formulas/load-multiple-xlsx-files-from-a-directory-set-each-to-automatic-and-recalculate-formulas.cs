using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Directory containing the XLSX files
        string folderPath = @"C:\ExcelFiles";

        // Get all .xlsx files in the directory
        string[] files = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);

        foreach (string filePath in files)
        {
            // Load the workbook from file
            Workbook workbook = new Workbook(filePath);

            // Set calculation mode to Automatic
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the workbook (overwrites the original file)
            workbook.Save(filePath);
        }

        Console.WriteLine("All workbooks have been set to Automatic calculation mode and refreshed.");
    }
}