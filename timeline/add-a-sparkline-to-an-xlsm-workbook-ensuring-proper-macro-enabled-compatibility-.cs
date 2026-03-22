using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace SparklineXlsmDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (XLSX by default, will be saved as XLSM later)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline (row 1, columns A to D)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            // Define the cell where the sparkline will be placed (E1)
            CellArea sparklineLocation = new CellArea
            {
                StartRow = 0,   // Row index is zero‑based (0 = first row)
                EndRow = 0,
                StartColumn = 4, // Column index is zero‑based (4 = column E)
                EndColumn = 4
            };

            // Add a sparkline group:
            // - Type: Line
            // - Data range: A1:D1
            // - isVertical: false (plot by columns)
            // - Location range: the cell area defined above
            int groupIndex = sheet.SparklineGroups.Add(
                SparklineType.Line,
                sheet.Name + "!A1:D1",
                false,
                sparklineLocation);

            // Retrieve the created group
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Add a sparkline to the group (the Add method also returns the index)
            // Parameters: data range, row index, column index
            group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

            // Optional: customize appearance (e.g., show high/low points)
            group.ShowHighPoint = true;
            group.ShowLowPoint = true;

            // Save the workbook as a macro‑enabled file (XLSM)
            // This ensures compatibility with Excel macros even though no VBA code is added.
            workbook.Save("SparklineDemo.xlsm", SaveFormat.Xlsm);

            Console.WriteLine("Workbook with sparkline saved as SparklineDemo.xlsm");
        }
    }
}