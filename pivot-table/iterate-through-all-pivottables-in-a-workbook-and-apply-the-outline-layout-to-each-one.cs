using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotOutline
{
    public class ApplyOutlineToAllPivotTables
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
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get pivot tables collection
                PivotTableCollection pivots = sheet.PivotTables;

                // Apply outline layout to each pivot table
                for (int i = 0; i < pivots.Count; i++)
                {
                    PivotTable pivot = pivots[i];
                    pivot.ShowInOutlineForm();   // layout in outline form
                    pivot.RefreshData();         // refresh data source
                    pivot.CalculateData();       // recalculate pivot
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }
}