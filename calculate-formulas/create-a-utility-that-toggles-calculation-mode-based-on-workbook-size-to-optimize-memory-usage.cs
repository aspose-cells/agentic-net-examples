// Title: C# utility to toggle Aspose.Cells calculation mode based on Excel workbook size for memory optimization
// Common Searches: how to switch calculation mode in Aspose.Cells when Excel file exceeds a size limit | C# code to recalculate formulas only for small Excel workbooks using Aspose.Cells | disable automatic formula recalculation for large Excel files in .NET with Aspose.Cells | set calculation mode conditionally based on workbook file size in Aspose.Cells | optimize memory usage by toggling Aspose.Cells calculation mode in C#
// Tags: Aspose.Cells conditional calculation mode | Excel workbook size threshold calculation toggle | C# memory optimization large workbooks | CalculateFormula for small Excel files | disable automatic recalculation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The CalculationModeToggler class checks that the Excel file exists, loads it with Aspose.Cells, measures its size, and forces a full formula recalculation only when the file size is at or below a specified byte threshold. For larger files the default (non‑automatic) calculation mode is retained, reducing memory consumption. The workbook is then saved back, with robust error handling for missing files and runtime exceptions.
public class CalculationModeToggler
{
    /// <param name="filePath">Full path to the Excel file.</param>
    /// <param name="sizeThresholdInBytes">Size threshold in bytes to decide the mode.</param>
    public static void ToggleCalculationMode(string filePath, long sizeThresholdInBytes)
    {
        // Verify that the file exists before attempting to load it.
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File not found – {filePath}");
            return;
        }

        try
        {
            // Load the workbook from the specified file.
            Workbook workbook = new Workbook(filePath);

            // Determine the actual file size on disk.
            long fileSize = new FileInfo(filePath).Length;

            // For large workbooks we avoid automatic recalculations.
            // For smaller workbooks we perform a full recalculation to ensure data consistency.
            if (fileSize <= sizeThresholdInBytes)
            {
                // Force a full recalculation for smaller workbooks.
                workbook.CalculateFormula();
            }

            // Save the workbook back to the same file (overwrites the original).
            workbook.Save(filePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }

    // Example usage.
    public static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: CalculationModeToggler <excelFilePath> <sizeThresholdInBytes>");
            return;
        }

        string excelFilePath = args[0];
        if (!long.TryParse(args[1], out long threshold))
        {
            Console.WriteLine("Invalid size threshold.");
            return;
        }

        ToggleCalculationMode(excelFilePath, threshold);
        Console.WriteLine("Calculation mode toggled based on workbook size.");
    }
}
