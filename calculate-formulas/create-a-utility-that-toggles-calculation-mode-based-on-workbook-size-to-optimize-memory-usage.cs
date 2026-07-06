using System;
using Aspose.Cells;

class WorkbookOptimizer
{
    // Optimizes a workbook by toggling calculation mode and memory setting based on total used cells.
    public static void Optimize(string inputPath, string outputPath, long cellThreshold)
    {
        // Load the workbook from the specified file.
        Workbook workbook = new Workbook(inputPath);

        // Estimate total used cells across all worksheets.
        long totalCells = 0;
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;
            // Approximate used area: (max row index + 1) * (max column index + 1)
            totalCells += (long)(cells.MaxDataRow + 1) * (cells.MaxDataColumn + 1);
        }

        // If the workbook exceeds the threshold, switch to manual calculation and file‑cache memory mode.
        // Otherwise, keep automatic calculation and normal memory mode.
        if (totalCells > cellThreshold)
        {
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;
            workbook.Settings.MemorySetting = MemorySetting.FileCache;
        }
        else
        {
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;
            workbook.Settings.MemorySetting = MemorySetting.Normal;
        }

        // Save the optimized workbook to the output path.
        workbook.Save(outputPath);
    }
}

class Program
{
    static void Main()
    {
        // Example usage:
        string inputFile = "input.xlsx";
        string outputFile = "output.xlsx";
        long cellCountThreshold = 100_000; // Adjust based on your performance criteria.

        WorkbookOptimizer.Optimize(inputFile, outputFile, cellCountThreshold);

        Console.WriteLine("Workbook optimization completed.");
    }
}