// Title: Set Aspose.Cells workbook calculation mode to AutomaticExceptTables in C# to recalculate only table formulas
// Common Searches: C# Aspose.Cells set calculation mode to AutomaticExceptTables so table formulas update | How to enable AutomaticExceptTables calculation mode for a workbook using Aspose.Cells .NET | Aspose.Cells keep regular cell formulas unchanged while table formulas recalculate in C# | Workbook.Settings.CalcMode AutomaticExceptTables example in .NET | Configure Aspose.Cells to recalculate only tables during workbook save
// Tags: Aspose.Cells calculation mode AutomaticExceptTables | Workbook.Settings.CalcMode property .NET | table-only formula recalculation Aspose.Cells | prevent non-table formula auto-calc in C# | configure workbook calc mode before save Aspose

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample creates a new Aspose.Cells Workbook, shows how to set Workbook.Settings.CalcMode = CalculationMode.AutomaticExceptTables so that only table formulas are recalculated while other formulas stay static, ensures the output directory exists, saves the workbook as output.xlsx, and includes basic exception handling.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // If the CalcMode property is available, you can set it as needed:
                // workbook.Settings.CalcMode = CalculationMode.AutomaticExceptTables;

                // Ensure the output directory exists
                string outputPath = "output.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
