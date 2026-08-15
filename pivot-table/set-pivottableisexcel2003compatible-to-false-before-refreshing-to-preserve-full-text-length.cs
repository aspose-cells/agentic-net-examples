// Title: Preserve Long Text in Aspose.Cells PivotTable by Turning Off Excel 2003 Compatibility (C#)
// Description: Demonstrates creating a workbook with descriptions longer than 255 characters, building a pivot table, setting PivotTable.IsExcel2003Compatible = false before RefreshData, and saving the file so the full text appears in the pivot results.
// Keywords: Aspose.Cells C# pivot table | IsExcel2003Compatible false | long text truncation Excel 2003 | RefreshData Aspose.Cells | CalculateData pivot | preserve string length pivot | .NET Excel API | pivot cache refresh | Excel 2003 compatibility flag | Aspose.Cells example
// Common Searches: Aspose.Cells keep full text in pivot table | disable Excel 2003 compatibility before RefreshData | pivot table string limit 255 characters C# | how to stop text truncation in Aspose.Cells pivot | set IsExcel2003Compatible property Aspose.Cells
// Developer Intent: The developer needs to disable Excel 2003 compatibility on a PivotTable so that long string values are not cut off when the pivot cache is refreshed.
// Use Cases: Generate product catalogs where descriptions exceed 255 characters and must appear completely in a pivot view. | Export detailed customer comments to Excel and retain the full comment text in pivot reports. | Build analytical dashboards that include verbose fields without losing data due to legacy limits.
// AI Prompts: Show C# code that creates an Aspose.Cells pivot table with long text fields and disables Excel 2003 compatibility. | Explain why PivotTable.IsExcel2003Compatible affects string truncation and how to avoid it in Aspose.Cells. | Provide a step‑by‑step guide to refresh a pivot cache after setting IsExcel2003Compatible = false.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates creating a workbook with descriptions longer than 255 characters, building a pivot table, setting PivotTable.IsExcel2003Compatible = false before RefreshData, and saving the file so the full text appears in the pivot results.
    public class PivotTableExcel2003CompatibilityDemo
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
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and add sample data
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Header row
            dataSheet.Cells["A1"].Value = "Product";
            dataSheet.Cells["B1"].Value = "Description";

            // Row 2 – short description
            dataSheet.Cells["A2"].Value = "Product1";
            dataSheet.Cells["B2"].Value = "Short description";

            // Row 3 – very long description (exceeds 255 characters)
            dataSheet.Cells["A3"].Value = "Product2";
            dataSheet.Cells["B3"].Value = new string('X', 300); // 300‑character string

            // Add a new worksheet to host the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Create the pivot table (source range A1:B3, destination cell A5)
            int pivotIndex = pivotSheet.PivotTables.Add("Data!A1:B3", "A5", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Add fields: Product as row, Description as data
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);    // Column 0 – Product
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);   // Column 1 – Description

            // Disable Excel 2003 compatibility to keep full text length
            pivotTable.IsExcel2003Compatible = false;

            // Refresh the pivot cache and calculate the results
            pivotTable.RefreshData();      // Correct API to refresh pivot cache
            pivotTable.CalculateData();

            // Save the workbook
            string outputPath = "PivotTableExcel2003CompatibilityDemo.xlsx";

            try
            {
                // Ensure the directory exists
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
