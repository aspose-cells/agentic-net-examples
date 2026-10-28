// Title: How to speed up Aspose.Cells .NET Workbook.CalculateFormula by optimizing custom function logic
// Common Searches: asp.net reduce Aspose.Cells CalculateFormula execution time for large worksheets | optimize custom function performance in Aspose.Cells C# | fast Excel formula recalculation using Aspose.Cells .NET after cell updates | benchmark Aspose.Cells formula engine and improve calculation speed | minimize latency of Workbook.CalculateFormula in C# applications
// Tags: Aspose.Cells formula engine performance tuning | C# workbook calculation optimization | custom function execution speed Aspose.Cells | Excel formula recalculation latency reduction | Aspose.Cells large dataset calculation efficiency

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, writes numeric values to cells A1 and A2, assigns a SUM formula to A3, executes Workbook.CalculateFormula, modifies A1, runs CalculateFormula again, and saves the workbook as Optimized.xlsx. It provides a baseline for measuring and improving formula recalculation performance in Aspose.Cells for .NET.
public class Program
{
    public static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate test data.
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);

            // Use the built‑in SUM function (custom functions require additional assemblies).
            sheet.Cells["A3"].Formula = "SUM(A1,A2)";

            // First calculation.
            workbook.CalculateFormula();

            // Modify a cell to trigger a second calculation.
            sheet.Cells["A1"].PutValue(15);

            // Second calculation.
            workbook.CalculateFormula();

            // Ensure the output directory exists.
            string outputPath = "Optimized.xlsx";
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
