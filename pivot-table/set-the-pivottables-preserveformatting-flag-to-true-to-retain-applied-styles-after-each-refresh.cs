// Title: Aspose.Cells C# – Enable PreserveFormatting on a PivotTable to Keep Styles After Refresh
// Description: Demonstrates how to load or create an Excel workbook, access its first PivotTable, set the PreserveFormatting flag to true, refresh and recalculate the data, and save the file so that all custom styles survive subsequent refreshes.
// Keywords: Aspose.Cells PivotTable PreserveFormatting | C# keep pivot styles after refresh | Aspose.Cells refresh pivot table | retain pivot formatting Aspose | PivotTable PreserveFormatting true
// Common Searches: Aspose.Cells keep pivot table formatting after refresh | Set PreserveFormatting property in Aspose.Cells C# | Refresh pivot table without losing style Aspose | How to preserve pivot table styles using Aspose.Cells
// Developer Intent: Turn on the PreserveFormatting flag for a PivotTable so its formatting is retained each time the table is refreshed.
// Use Cases: Generate a daily report: create a workbook, add a pivot table, enable PreserveFormatting, refresh data, and save with consistent styling. | Update an existing analysis file: load a workbook, activate PreserveFormatting on its pivot tables, refresh source data, and preserve custom number formats and colors. | Automate data pipelines: programmatically recalculate pivot tables across multiple workbooks while ensuring predefined fonts, borders, and conditional formats remain unchanged.
// AI Prompts: Show C# code to set PivotTable.PreserveFormatting = true with Aspose.Cells and refresh the table. | Provide an Aspose.Cells example that creates a pivot table, applies a style, enables PreserveFormatting, and saves the workbook. | Explain how PreserveFormatting interacts with RefreshData and CalculateData in Aspose.Cells PivotTables.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to load or create an Excel workbook, access its first PivotTable, set the PreserveFormatting flag to true, refresh and recalculate the data, and save the file so that all custom styles survive subsequent refreshes.
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
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        public static void Run()
        {
            string inputPath = "source.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists; create a sample workbook if it does not.
            if (!File.Exists(inputPath))
            {
                try
                {
                    CreateSampleWorkbook(inputPath);
                    Console.WriteLine($"Sample workbook created at '{inputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Failed to create sample workbook: " + ex.Message);
                    return;
                }
            }

            // Load the workbook.
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to load workbook: " + ex.Message);
                return;
            }

            // Get the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Verify that a pivot table exists.
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Access the first pivot table.
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Preserve formatting when the pivot table is refreshed.
            pivotTable.PreserveFormatting = true;

            try
            {
                // Refresh data from the source range.
                pivotTable.RefreshData();

                // Recalculate the pivot table.
                pivotTable.CalculateData();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while refreshing/recalculating pivot table: " + ex.Message);
                return;
            }

            // Save the modified workbook.
            try
            {
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to save workbook: " + ex.Message);
            }
        }

        private static void CreateSampleWorkbook(string path)
        {
            // Create a new workbook with sample data.
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data.
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Amount");
            ws.Cells["A2"].PutValue("Food");
            ws.Cells["B2"].PutValue(120);
            ws.Cells["A3"].PutValue("Transport");
            ws.Cells["B3"].PutValue(80);
            ws.Cells["A4"].PutValue("Food");
            ws.Cells["B4"].PutValue(150);
            ws.Cells["A5"].PutValue("Transport");
            ws.Cells["B5"].PutValue(70);

            // Define the source data range for the pivot table.
            string sourceData = "Sheet1!A1:B5";

            // Add a pivot table.
            string destinationCell = "A7";
            int pivotIndex = ws.PivotTables.Add(sourceData, "PivotTable1", destinationCell, true);
            PivotTable pt = ws.PivotTables[pivotIndex];

            // Note: Adding fields by name may require additional code depending on the Aspose.Cells version.
            // For demonstration purposes, the pivot table is left with default fields.

            // Save the sample workbook.
            wb.Save(path, SaveFormat.Xlsx);
        }
    }
}
