using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCircularReferenceReport
{
    // Custom monitor to capture circular reference information
    public class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        // Shared list to store report entries
        public static List<string> CircularInfo { get; } = new List<string>();

        public override bool OnCircular(IEnumerator circularCellsData)
        {
            // Iterate through all cells involved in the circular reference
            while (circularCellsData.MoveNext())
            {
                // Each item is a CalculationCell
                var calcCell = circularCellsData.Current as CalculationCell;
                if (calcCell != null)
                {
                    // Retrieve the actual cell object using row/column indexes
                    // Aspose.Cells versions may expose Row/Column or RowIndex/ColumnIndex
                    int row = 0, column = 0;
                    if (calcCell.GetType().GetProperty("Row") != null)
                    {
                        row = (int)calcCell.GetType().GetProperty("Row")!.GetValue(calcCell)!;
                        column = (int)calcCell.GetType().GetProperty("Column")!.GetValue(calcCell)!;
                    }
                    else
                    {
                        row = (int)calcCell.GetType().GetProperty("RowIndex")!.GetValue(calcCell)!;
                        column = (int)calcCell.GetType().GetProperty("ColumnIndex")!.GetValue(calcCell)!;
                    }

                    Cell cell = calcCell.Worksheet.Cells[row, column];
                    // Store address and formula for reporting
                    CircularInfo.Add($"{cell.Name}: {cell.Formula}");
                }
            }
            // Continue calculation (return false to abort)
            return true;
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Set up some circular references for demonstration
                cells["A1"].Formula = "=B1";
                cells["B1"].Formula = "=A1";
                cells["C1"].Formula = "=C1+1"; // Self‑referencing circular formula

                // Configure calculation options with the custom monitor
                CalculationOptions options = new CalculationOptions
                {
                    CalculationMonitor = new CircularReferenceMonitor()
                };

                // Perform formula calculation; the monitor will capture circular cells
                workbook.CalculateFormula(options);

                // Output report to console
                Console.WriteLine("Circular Reference Report:");
                foreach (string entry in CircularReferenceMonitor.CircularInfo)
                {
                    Console.WriteLine(entry);
                }

                // Create a new worksheet to store the report inside the workbook
                Worksheet reportSheet = workbook.Worksheets.Add("CircularReport");
                reportSheet.Cells[0, 0].PutValue("Cell");
                reportSheet.Cells[0, 1].PutValue("Formula");

                int reportRow = 1;
                foreach (string entry in CircularReferenceMonitor.CircularInfo)
                {
                    // Split the stored string into address and formula parts
                    string[] parts = entry.Split(new[] { ':' }, 2);
                    if (parts.Length == 2)
                    {
                        reportSheet.Cells[reportRow, 0].PutValue(parts[0].Trim());
                        reportSheet.Cells[reportRow, 1].PutValue(parts[1].Trim());
                        reportRow++;
                    }
                }

                // Ensure the output directory exists
                string outputPath = "CircularReferenceReport.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath))!;
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with the report sheet
                workbook.Save(outputPath);
                Console.WriteLine($"Report saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}