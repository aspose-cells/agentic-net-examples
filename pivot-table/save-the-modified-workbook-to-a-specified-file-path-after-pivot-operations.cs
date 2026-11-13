// Title: Refresh all pivot tables and save the workbook to a custom file path using Aspose.Cells for .NET (C#)
// Common Searches: C# code to refresh every pivot table in an Excel workbook and save to a new file with Aspose.Cells | Aspose.Cells .NET save modified workbook to custom directory after pivot refresh | How to ensure output folder exists before calling Workbook.Save in Aspose.Cells | RefreshData and CalculateData for all pivot tables then export as .xlsx using Aspose.Cells | Load existing .xlsx, refresh pivots, and write to another location with Aspose.Cells for .NET
// Tags: refresh all pivot tables Aspose.Cells API | Workbook.Save to specific location C# | SaveFormat.Xlsx usage Aspose.Cells | create output directory programmatically C# | iterate worksheets to access pivot tables Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable class

namespace PivotWorkbookExample
{
    // The example loads an existing Excel workbook, iterates through each worksheet to refresh and calculate all pivot tables, ensures the target output directory exists, and then saves the updated workbook to a specified path in XLSX format using Aspose.Cells for .NET, with error handling.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths
            string inputFilePath = @"C:\Data\SourceWorkbook.xlsx";
            string outputFilePath = @"C:\Data\ModifiedWorkbook.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputFilePath))
            {
                Console.WriteLine($"Input file not found: {inputFilePath}");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputFilePath);

                // Refresh all pivot tables in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (PivotTable pivotTable in sheet.PivotTables)
                    {
                        pivotTable.RefreshData();
                        pivotTable.CalculateData();
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputFilePath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook to the specified file path
                workbook.Save(outputFilePath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to: {outputFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
