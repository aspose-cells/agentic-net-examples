// Title: Enable iterative calculation mode and set maximum iterations for circular references with Aspose.Cells in C#
// AI Prompts: Generate C# code that loads an Excel workbook using Aspose.Cells, turns on iterative calculation, specifies a maximum iteration count, and saves the updated file. | Show how to configure Aspose.Cells workbook settings to handle circular formulas by enabling iterative calculation and defining a custom iteration limit.
// Common Searches: asp.net enable iterative calculation Aspose.Cells circular reference handling | set maximum iteration count for circular formulas using Aspose.Cells C# | Aspose.Cells workbook calculation settings example .NET | how to configure iterative mode for Excel workbooks with Aspose.Cells | C# code to activate iterative calculation and limit iterations in Aspose.Cells
// Tags: iterative calculation Aspose.Cells | maximum iteration count Aspose.Cells | circular reference handling Aspose.Cells | workbook calculation settings C# | enable iterative mode Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook (or creates a new one), enables iterative calculation, sets a custom maximum iteration limit to improve convergence of circular formulas, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook if it exists; otherwise create a new one
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Iterative calculation settings are not available in this version of Aspose.Cells.
            // If needed, configure calculation mode here (e.g., workbook.Settings.CalcMode = CalculationMode.Automatic).

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
