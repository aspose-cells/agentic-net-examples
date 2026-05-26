using System;
using System.IO;
using Aspose.Cells;

public class CalculationModeToggler
{
    /// <summary>
    /// Toggles the workbook's calculation mode based on the file size.
    /// If the file size is greater than <paramref name="sizeThresholdBytes"/>, the mode is set to Manual;
    /// otherwise it is set to Automatic.
    /// </summary>
    /// <param name="inputPath">Path to the source workbook.</param>
    /// <param name="outputPath">Path where the modified workbook will be saved.</param>
    /// <param name="sizeThresholdBytes">Size threshold in bytes.</param>
    public static void ToggleCalculationMode(string inputPath, string outputPath, long sizeThresholdBytes)
    {
        // Load the workbook from the specified file.
        Workbook workbook = new Workbook(inputPath);

        // Determine the file size.
        long fileSize = new FileInfo(inputPath).Length;

        // Choose calculation mode based on size.
        if (fileSize > sizeThresholdBytes)
        {
            // Large file – switch to Manual calculation to improve performance.
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;
        }
        else
        {
            // Small file – use Automatic calculation.
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;
        }

        // Save the workbook with the updated setting.
        workbook.Save(outputPath);
    }

    // Example usage
    public static void Main()
    {
        string sourceFile = "input.xlsx";
        string resultFile = "output.xlsx";
        long threshold = 5 * 1024 * 1024; // 5 MB

        ToggleCalculationMode(sourceFile, resultFile, threshold);

        Console.WriteLine("Calculation mode toggled based on file size and saved to " + resultFile);
    }
}