// Title: How to keep user‑applied cell formatting when refreshing a PivotTable using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an existing .xlsx workbook with Aspose.Cells, sets PivotTable.PreserveCellFormattingOnUpdate to true, refreshes the pivot data, and saves the file. | Show how to update the first PivotTable in a worksheet while preserving all custom cell styles after calling RefreshData and CalculateData with Aspose.Cells. | Demonstrate enabling PreserveCellFormattingOnUpdate for a PivotTable, then refreshing and recalculating the pivot to maintain formatting in a .NET application.
// Common Searches: Aspose.Cells C# preserve pivot table formatting after RefreshData | How to retain cell styles in a PivotTable when updating data source with Aspose.Cells | Set PreserveCellFormattingOnUpdate property in Aspose.Cells .NET example | Refresh PivotTable without losing custom formatting using Aspose.Cells for .NET
// Tags: Aspose.Cells pivot formatting preservation | PivotTable PreserveCellFormattingOnUpdate usage | Aspose.Cells refresh pivot data | Aspose.Cells recalculate pivot | Aspose.Cells .NET pivot table update

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example loads 'source.xlsx', checks for a PivotTable, enables PreserveCellFormattingOnUpdate to keep user‑applied styles, refreshes and recalculates the pivot, and saves the result as 'output.xlsx' using Aspose.Cells for .NET.
    public class PivotTablePreserveFormattingDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            string inputPath = "source.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook containing the pivot table
                Workbook workbook = new Workbook(inputPath);
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one pivot table
                if (worksheet.PivotTables.Count > 0)
                {
                    PivotTable pivotTable = worksheet.PivotTables[0];

                    // Preserve cell formatting when the pivot table is refreshed
                    pivotTable.PreserveCellFormattingOnUpdate = true;

                    // Refresh the pivot table data source
                    pivotTable.RefreshData();

                    // Recalculate the pivot table to apply refreshed data
                    pivotTable.CalculateData();
                }
                else
                {
                    Console.WriteLine("No pivot tables found in the worksheet.");
                }

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }
}
