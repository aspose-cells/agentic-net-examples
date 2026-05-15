using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";
        // Path where the processed workbook will be saved
        string outputPath = "output.xlsx";

        // Open a read‑only file stream
        using (FileStream stream = new FileStream(sourcePath, FileMode.Open, FileAccess.Read))
        {
            // Create LoadOptions and ensure formulas are parsed when the file is opened
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.ParsingFormulaOnOpen = true; // default, set explicitly for clarity

            // Load the workbook from the stream using the specified options
            Workbook workbook = new Workbook(stream, loadOptions);

            // Enable automatic formula calculation mode
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // Calculate all formulas immediately (optional, ensures values are up‑to‑date)
            workbook.CalculateFormula();

            // Save the workbook to a new file
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}