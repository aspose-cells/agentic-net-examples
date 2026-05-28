using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class RefreshNestedPivotTables
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

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing the parent and child pivot tables
            Workbook workbook = new Workbook(inputPath);

            // Assume the parent pivot table is on the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
                return;
            }

            // Get the parent pivot table
            PivotTable parentPivot = worksheet.PivotTables[0];

            // Refresh the parent pivot table to update its cache
            parentPivot.RefreshData();
            parentPivot.CalculateData();

            // Retrieve dependent (child) pivot tables
            PivotTable[] childPivots = parentPivot.GetDependentPivotTables();

            // Refresh each child pivot table to synchronize with the updated parent cache
            foreach (PivotTable child in childPivots)
            {
                child.RefreshData();
                child.CalculateData();
            }

            // Save the workbook with refreshed pivot tables
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}