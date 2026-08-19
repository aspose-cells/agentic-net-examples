// Title: Clear AutoFilter on column AJ and show all rows with Aspose.Cells for .NET
// Description: Loads an Excel file, ensures the AutoFilter range includes column AJ, calls worksheet.AutoFilter.ShowAll() and Refresh() to remove filter criteria, then saves the workbook with every row visible.
// Keywords: Aspose.Cells | .NET | Clear AutoFilter | ShowAll | AutoFilter.Refresh | column AJ | remove Excel filter | display all rows | reset worksheet filter | Excel automation
// Common Searches: Aspose.Cells clear filter column AJ | How to show all rows after AutoFilter in C# | Reset Excel AutoFilter programmatically | Aspose.Cells ShowAll example | Remove specific column filter with Aspose.Cells
// Developer Intent: Remove any AutoFilter criteria applied to column AJ and make every worksheet row visible.
// Use Cases: Prepare a report workbook for distribution by clearing hidden rows caused by filters. | Ensure data export includes the full dataset after temporary filtering. | Programmatically reset worksheet filters before performing bulk calculations or chart generation.
// AI Prompts: Generate C# code using Aspose.Cells to clear the AutoFilter on column AJ while keeping other filters intact. | Explain how worksheet.AutoFilter.ShowAll() and Refresh() work together when resetting filters. | Show how to verify or set an AutoFilter range before calling ShowAll in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an Excel file, ensures the AutoFilter range includes column AJ, calls worksheet.AutoFilter.ShowAll() and Refresh() to remove filter criteria, then saves the workbook with every row visible.
    public class ClearAutoFilterOnColumnAJ
    {
        public static void Run()
        {
            string inputPath = "InputWorkbook.xlsx";
            string outputPath = "OutputWorkbook.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure an AutoFilter range is defined (optional)
                if (string.IsNullOrEmpty(worksheet.AutoFilter?.Range))
                {
                    worksheet.AutoFilter.Range = "A1:AJ100";
                }

                // Clear all filter criteria (including column AJ)
                worksheet.AutoFilter.ShowAll();

                // Refresh to apply changes (optional after ShowAll)
                worksheet.AutoFilter.Refresh();

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ClearAutoFilterOnColumnAJ.Run();
        }
    }
}
