// Title: Recalculate all Excel formulas in a workbook using Aspose.Cells Workbook.CalculateFormula (C#)
// Common Searches: c# how to recalculate all formulas in an Excel file using Aspose.Cells | Aspose.Cells Workbook.CalculateFormula example with loading and saving workbook | using Aspose.Cells to force formula evaluation after modifying cells in C# | Aspose.Cells custom calculation monitor not available in current version C# | recalculate Excel workbook formulas and save output with Aspose.Cells C#
// Tags: Aspose.Cells Workbook.CalculateFormula usage | Excel formula recalculation C# Aspose.Cells | load or create workbook before formula evaluation Aspose.Cells | save workbook after recalculating formulas Aspose.Cells | custom calculation monitor interface missing Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file (or creates a new workbook if the file is missing), recalculates every formula using Workbook.CalculateFormula, and then saves the workbook. It also notes that the custom calculation monitor interface is unavailable in the referenced Aspose.Cells version.
public class Program
{
    public static void Main()
    {
        Workbook workbook = null;

        try
        {
            // Load an existing workbook if the file exists; otherwise create a new one.
            string inputPath = "input.xlsx";
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
                Console.WriteLine($"Workbook loaded from \"{inputPath}\".");
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook
                Console.WriteLine("Input file not found. Created a new workbook.");
            }

            // Recalculate all formulas.
            workbook.CalculateFormula();
            Console.WriteLine("Formula recalculation completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during workbook processing: {ex.Message}");
            return;
        }

        try
        {
            // Save the workbook.
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}

// Note: Custom calculation monitor removed because the required interface
// (ICalculationMonitor) may not be available in the referenced Aspose.Cells version.
