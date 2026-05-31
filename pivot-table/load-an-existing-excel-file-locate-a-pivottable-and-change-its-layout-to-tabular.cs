using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace MyApp
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "InputWithPivot.xlsx";
                const string outputPath = "OutputTabularPivot.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Locate the first PivotTable
                PivotTable pivotTable = null;
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.PivotTables.Count > 0)
                    {
                        pivotTable = sheet.PivotTables[0];
                        break;
                    }
                }

                // Change layout to Tabular form if a PivotTable exists
                if (pivotTable != null)
                {
                    pivotTable.ShowInTabularForm();   // Layout the PivotTable in tabular form
                    pivotTable.RefreshData();         // Refresh data from the source
                    pivotTable.CalculateData();       // Recalculate the PivotTable
                }
                else
                {
                    Console.WriteLine("No PivotTable found in the workbook.");
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}