// Title: PreserveFormatting for PivotTable in Aspose.Cells .NET – Keep Custom Styles After Refresh
// Description: Loads a workbook, sets PivotTable.PreserveFormatting = true, applies a bold Arial style with a light‑blue background via FormatAll, refreshes and recalculates the pivot, then saves the file, ensuring user‑defined formatting persists.
// Keywords: Aspose.Cells | C# | PivotTable PreserveFormatting | RefreshData | FormatAll | custom pivot style | keep formatting after refresh | Excel automation | pivot table styling | Aspose.Cells .NET example
// Common Searches: Aspose.Cells keep pivot formatting after refresh | Set PreserveFormatting on PivotTable C# | Apply style to entire PivotTable Aspose.Cells | RefreshData without losing pivot styles | PivotTable FormatAll example
// Developer Intent: Enable the PreserveFormatting flag on a PivotTable so that any user‑applied formatting remains intact when the table is refreshed or recalculated.
// Use Cases: Automated reporting where branding colors and fonts must survive data refreshes. | Batch processing of workbooks to apply a corporate style to all pivot tables and retain it across updates. | Building a .NET service that refreshes pivot data daily without overwriting custom formatting.
// AI Prompts: Generate C# code using Aspose.Cells to set PreserveFormatting on a PivotTable, apply a bold Arial style with a light blue background, refresh the data, and save the workbook. | Explain how PreserveFormatting interacts with RefreshData and CalculateData in Aspose.Cells PivotTable processing. | Provide a step‑by‑step tutorial for styling a PivotTable programmatically and preserving the style after multiple refresh cycles in .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using System.Drawing;

namespace AsposeCellsExamples
{
    // Loads a workbook, sets PivotTable.PreserveFormatting = true, applies a bold Arial style with a light‑blue background via FormatAll, refreshes and recalculates the pivot, then saves the file, ensuring user‑defined formatting persists.
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
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file '{inputPath}' not found.");

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one pivot table
                if (worksheet.PivotTables.Count == 0)
                    throw new InvalidOperationException("No pivot tables found in the worksheet.");

                // Get the first pivot table
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Preserve formatting when the pivot table is refreshed
                pivotTable.PreserveFormatting = true;

                // Create a style for the pivot table data area
                Style style = workbook.CreateStyle();
                style.Font.Name = "Arial";
                style.Font.Size = 10;
                style.Font.IsBold = true;
                style.ForegroundColor = Color.LightBlue;
                style.Pattern = BackgroundType.Solid;

                // Apply the style to the entire pivot table
                pivotTable.FormatAll(style);

                // Refresh the pivot table data
                pivotTable.RefreshData();

                // Recalculate the pivot table after refresh
                pivotTable.CalculateData();

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any errors that occur during processing
                Console.WriteLine($"Processing error: {ex.Message}");
                throw;
            }
        }
    }
}
