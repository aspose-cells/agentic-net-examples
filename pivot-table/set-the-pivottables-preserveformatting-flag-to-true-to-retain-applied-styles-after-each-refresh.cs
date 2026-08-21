// Title: How to retain pivot table cell formatting after refresh using Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing .xlsx workbook, set PivotTable.PreserveCellFormattingOnUpdate = true, call RefreshData and CalculateData, then save the file with Aspose.Cells in C#. | Enable formatting preservation for a pivot table, refresh its source data, recalculate the pivot, and write the updated workbook to a new Excel file using Aspose.Cells for .NET. | Programmatically prevent style loss when updating a PivotTable by toggling PreserveCellFormattingOnUpdate, invoking RefreshData and CalculateData, and saving the workbook in C#.
// Common Searches: Aspose.Cells C# keep pivot table styles after RefreshData | Set PreserveCellFormattingOnUpdate flag for pivot table in .NET | C# example to save Excel workbook with pivot table formatting preserved using Aspose | How to avoid losing cell formatting when refreshing a pivot table with Aspose.Cells | Refresh pivot table without resetting custom formatting Aspose.Cells .NET
// Tags: Aspose.Cells PreserveCellFormattingOnUpdate | C# refresh pivot table retain formatting | Aspose.Cells pivot table formatting preservation | Save Excel workbook with unchanged pivot styles | PivotTable RefreshData CalculateData Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Loads a workbook, sets PreserveCellFormattingOnUpdate on the first PivotTable, refreshes and recalculates the pivot, and saves the workbook to a new XLSX file.
    public class PivotTablePreserveFormattingDemo
    {
        public static void Run()
        {
            try
            {
                const string inputPath = "source.xlsx";
                const string outputPath = "output.xlsx";

                // Ensure the source file exists
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException($"Input file '{inputPath}' not found.");
                }

                // Load the workbook containing a PivotTable
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Verify a PivotTable exists
                if (worksheet.PivotTables.Count == 0)
                {
                    throw new InvalidOperationException("No PivotTable found in the first worksheet.");
                }

                // Access the first PivotTable
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Preserve cell formatting during refresh
                pivotTable.PreserveCellFormattingOnUpdate = true;

                // Refresh the PivotTable data from its source
                pivotTable.RefreshData();

                // Recalculate the PivotTable data
                pivotTable.CalculateData();

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            PivotTablePreserveFormattingDemo.Run();
        }
    }
}
