// Title: How to configure a custom calculation engine timeout for formula evaluation in Aspose.Cells for .NET (C#)
// Common Searches: c# set calculation engine timeout in Aspose.Cells to avoid hanging | aspnet prevent long‑running Excel formulas using Aspose.Cells timeout | example of customizing formula calculation timeout with Aspose.Cells .NET | adjust Aspose.Cells formula recalculation timeout programmatically | Aspose.Cells workbook.CalculateFormula with custom timeout C#
// Tags: Aspose.Cells calculation engine timeout | C# custom formula evaluation timeout | prevent long-running Excel formulas Aspose.Cells | Aspose.Cells workbook recalculate with timeout | set formula calculation timeout .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, verifies the input file, recalculates all formulas using the default calculation engine timeout, and saves the result. It includes basic exception handling and demonstrates where you can set a custom timeout via the WorkbookSettings.CalcEngineTimeout property to prevent long‑running formulas from hanging the application.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Recalculate all formulas (default timeout is used)
            workbook.CalculateFormula();

            // Save the workbook after calculation
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors (e.g., loading, calculation, saving)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
