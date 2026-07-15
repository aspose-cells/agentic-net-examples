using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class ApplyCompactLayoutToPivotTable
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
            const string inputPath = "ExistingPivotTable.xlsx";
            const string outputPath = "ExistingPivotTable_CompactLayout.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook containing the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (worksheet.PivotTables.Count == 0)
                throw new InvalidOperationException("No PivotTable found in the worksheet.");

            // Get the first PivotTable
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Apply the Compact layout to the PivotTable
            pivotTable.ShowInCompactForm();

            // Refresh and recalculate the PivotTable to reflect changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
    }
}