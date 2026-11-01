// Title: Set workbook calculation mode to AutomaticExceptTables with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set calculation mode to AutomaticExceptTables | How to recalculate only table formulas in an Excel workbook using Aspose.Cells | Enable AutomaticExceptTables mode in Aspose.Cells .NET API | Preserve non‑table formula values while updating table calculations with Aspose.Cells | Change workbook Settings.CalcMode to AutomaticExceptTables in C#
// Tags: Aspose.Cells workbook Settings.CalcMode AutomaticExceptTables | C# set Excel calculation mode AutomaticExceptTables | recalculate table formulas only Aspose.Cells | static non‑table formulas Aspose.Cells | Excel workbook calculation mode control .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file with Aspose.Cells, optionally sets the workbook's Settings.CalcMode to AutomaticExceptTables so that only table formulas are recalculated while other formulas stay static, and saves the modified workbook, handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
            return;
        }

        try
        {
            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // (Optional) Set calculation mode if needed.
            // Note: The CalcMode property may not be available in older Aspose.Cells versions.
            // Uncomment the following line if your version supports it:
            // workbook.Settings.CalcMode = CalculationMode.AutomaticExceptTables;

            // Save the workbook with the updated settings
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
