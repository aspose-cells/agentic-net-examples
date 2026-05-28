using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using System.Drawing;

namespace AsposeCellsExamples
{
    public class PivotTablePreserveFormattingDemo
    {
        public static void Main()
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
            const string inputPath = "source.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the source workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook containing the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the first pivot table in the worksheet
            if (worksheet.PivotTables.Count == 0)
                throw new InvalidOperationException("No pivot tables found in the worksheet.");

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

            // Refresh data and recalculate the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}