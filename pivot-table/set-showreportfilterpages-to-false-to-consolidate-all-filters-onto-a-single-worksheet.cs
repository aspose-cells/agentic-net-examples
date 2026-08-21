// Title: Consolidate Pivot Table Report Filters onto One Worksheet (ShowReportFilterPages = false) – Aspose.Cells for .NET
// Description: Loads a workbook, finds the first pivot table, uses reflection to set its ShowReportFilterPages property to false (when supported), and saves the file so all report‑filter pages are merged into a single sheet.
// Keywords: Aspose.Cells ShowReportFilterPages false | pivot table consolidate filters .NET | disable pivot report filter pages | Aspose.Cells reflection property | single worksheet pivot filters | C# Aspose.Cells pivot table settings
// Common Searches: how to turn off ShowReportFilterPages in Aspose.Cells | merge pivot report filter pages into one sheet C# | set ShowReportFilterPages property via reflection | remove extra pivot filter worksheets Aspose.Cells | consolidate pivot table filters .NET
// Developer Intent: Set ShowReportFilterPages to false so a pivot table’s report filters are kept on a single worksheet instead of generating separate pages.
// Use Cases: Update an existing workbook to prevent pivot report filters from creating extra worksheets. | Maintain compatibility across Aspose.Cells versions by checking for the ShowReportFilterPages property at runtime. | Provide clear console messages when the workbook is missing, no pivot tables exist, or the property is unavailable.
// AI Prompts: Generate C# code that disables ShowReportFilterPages for every pivot table in a workbook, with version‑safe reflection handling. | Explain step‑by‑step how to use reflection to modify the ShowReportFilterPages property of a PivotTable in Aspose.Cells. | Create a reusable method that consolidates pivot report filter pages onto one sheet and logs detailed status messages for missing files or tables.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Loads a workbook, finds the first pivot table, uses reflection to set its ShowReportFilterPages property to false (when supported), and saves the file so all report‑filter pages are merged into a single sheet.
    public class ConsolidatePivotReportFilters
    {
        public static void Run()
        {
            const string inputPath = "source.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the source file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook containing the pivot table
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Verify that at least one pivot table exists
                if (worksheet.PivotTables.Count > 0)
                {
                    PivotTable pivotTable = worksheet.PivotTables[0];

                    // Consolidate report filter pages onto a single sheet.
                    // Use reflection to set ShowReportFilterPages if the property exists in the current Aspose.Cells version.
                    var prop = pivotTable.GetType().GetProperty("ShowReportFilterPages");
                    if (prop != null && prop.CanWrite)
                    {
                        prop.SetValue(pivotTable, false);
                    }
                    else
                    {
                        Console.WriteLine("ShowReportFilterPages property is not available in this Aspose.Cells version.");
                    }
                }
                else
                {
                    Console.WriteLine("No pivot tables found in the worksheet.");
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ConsolidatePivotReportFilters.Run();
        }
    }
}
