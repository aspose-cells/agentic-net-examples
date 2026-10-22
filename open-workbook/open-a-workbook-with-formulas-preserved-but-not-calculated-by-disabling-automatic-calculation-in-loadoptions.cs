// Title: Load an Excel workbook in C# with Aspose.Cells, preserve formulas and disable automatic calculation using LoadOptions
// Common Searches: c# asp.net load excel file without recalculating formulas using aspose.cells | how to disable automatic calculation when opening a workbook with Aspose.Cells LoadOptions | preserve cell formulas on load with Aspose.Cells C# example | using LoadOptions to turn off formula evaluation in Aspose.Cells C# | open xlsx in C# with Aspose.Cells while keeping formulas unchanged
// Tags: load workbook without evaluating formulas Aspose.Cells | disable automatic calculation via LoadOptions C# | preserve formulas during workbook load Aspose.Cells | Aspose.Cells LoadOptions turn off calculation engine | open Excel file without formula recalculation C#

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates loading an Excel workbook with Aspose.Cells in C# while keeping all formulas intact and preventing any automatic recalculation by configuring LoadOptions to disable the calculation engine, then saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook with default options (formulas are preserved).
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook; formulas remain unchanged and are not recalculated.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
