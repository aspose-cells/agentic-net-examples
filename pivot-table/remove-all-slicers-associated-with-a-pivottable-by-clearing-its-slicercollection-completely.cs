using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsExamples
{
    public class RemoveAllSlicersFromPivotDemo
    {
        // Entry point for the console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "InputWithPivot.xlsx";
            const string outputPath = "OutputWithoutSlicers.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one PivotTable on the sheet
                PivotTableCollection pivotTables = worksheet.PivotTables;
                if (pivotTables.Count > 0)
                {
                    // Get the slicer collection for the worksheet
                    SlicerCollection slicers = worksheet.Slicers;

                    // Clear all slicers associated with the PivotTable(s)
                    slicers.Clear();
                }
                else
                {
                    Console.WriteLine("No PivotTables found on the worksheet.");
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved without slicers to: {outputPath}");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine($"File not found: {fnfEx.FileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}