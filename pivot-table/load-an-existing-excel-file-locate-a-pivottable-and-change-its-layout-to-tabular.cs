using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotTabularDemo
{
    class Program
    {
        static void Main()
        {
            const string inputPath = "InputWithPivot.xlsx";
            const string outputPath = "OutputTabularPivot.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                PivotTable pivotTable = null;
                Worksheet pivotSheet = null;

                // Locate the first PivotTable in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.PivotTables.Count > 0)
                    {
                        pivotTable = sheet.PivotTables[0];
                        pivotSheet = sheet;
                        break;
                    }
                }

                if (pivotTable == null || pivotSheet == null)
                {
                    Console.WriteLine("No PivotTable found in the workbook.");
                    return;
                }

                // Change the layout of the found PivotTable to Tabular form
                pivotTable.ShowInTabularForm();

                // Refresh the worksheet to apply the layout change
                pivotSheet.RefreshPivotTables();

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"PivotTable layout changed to Tabular and workbook saved as '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}