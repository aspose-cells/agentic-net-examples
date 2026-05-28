using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace Demo
{
    public class PreservePivotFormattingDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input file not found: {inputPath}");

                // Load the workbook containing the PivotTable
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure a PivotTable exists
                if (worksheet.PivotTables.Count == 0)
                    throw new InvalidOperationException("No PivotTable found in the first worksheet.");

                // Access the first PivotTable
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Preserve formatting when the PivotTable is refreshed
                pivotTable.PreserveFormatting = true;

                // Create a custom style (currency format with background color)
                Style style = workbook.CreateStyle();
                style.Custom = "$#,##0.00";
                style.ForegroundColor = Color.LightBlue;
                style.Pattern = BackgroundType.Solid;

                // Apply the style to the data body range of the PivotTable
                pivotTable.Format(pivotTable.DataBodyRange, style);

                // Refresh data source and recalculate
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}