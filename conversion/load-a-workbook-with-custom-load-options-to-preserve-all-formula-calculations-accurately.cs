// Title: Load an Excel workbook with Aspose.Cells in C# while preserving original formula results
// Common Searches: C# Aspose.Cells load workbook without recalculating formulas | how to keep original formula values when opening an XLSX with Aspose.Cells | prevent automatic calculation on workbook load using Aspose.Cells .NET | preserve formula results while saving Excel file with Aspose.Cells | custom load options to retain formula calculations in Aspose.Cells C#
// Tags: load workbook preserve formulas Aspose.Cells | manual calculation mode Aspose.Cells C# | excel file load without recalculation .NET | save workbook keep formula values Aspose.Cells | custom load options formula retention Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example verifies the input XLSX file, loads it into an Aspose.Cells Workbook using default options that retain formulas and their calculated values, optionally switches the calculation mode to manual to stop automatic recalculation, ensures the output directory exists, saves the workbook, and handles any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook (default options preserve formulas and values)
            Workbook workbook = new Workbook(inputPath);

            // Optional: prevent automatic recalculation if you want to keep original values
            // Note: In some Aspose.Cells versions the property is CalcMode; if unavailable, this step can be omitted.
            // workbook.Settings.CalcMode = CalculationMode.Manual;

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
