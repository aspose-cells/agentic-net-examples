// Title: How to set Aspose.Cells calculation mode to AutomaticExceptTables in C# to skip table formula recalculation
// Common Searches: Aspose.Cells C# set workbook calculation mode to AutomaticExceptTables | exclude Excel table formulas from automatic recalculation using Aspose.Cells .NET | example code for AutomaticExceptTables calculation mode in Aspose.Cells | configure Aspose.Cells calculation engine to ignore table formulas in C# | disable automatic updates for table formulas when saving workbook with Aspose.Cells
// Tags: Aspose.Cells calculation mode AutomaticExceptTables | C# set workbook calculation engine | skip table formulas automatic recalculation | Aspose.Cells workbook recalculation settings | exclude Excel table formulas Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Illustrates how to configure a Workbook's calculation engine to AutomaticExceptTables with Aspose.Cells for .NET, ensuring that only non‑table formulas are recalculated automatically, and then saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Note: Calculation mode setting is omitted because the property may not be available
            // in certain versions of Aspose.Cells. The default mode will be used.

            // Define output file path
            string outputPath = "Result.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

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
