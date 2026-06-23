using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCopyRowsAndUpdateChart
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet (source sheet)
            Workbook workbook = new Workbook();
            Worksheet srcSheet = workbook.Worksheets[0];
            srcSheet.Name = "Source";

            // Populate sample data in rows 2‑20 (Excel rows, zero‑based indices 1‑19)
            for (int row = 1; row <= 19; row++)          // rows 2‑20
            {
                srcSheet.Cells[row, 0].PutValue($"Item {row}");
                srcSheet.Cells[row, 1].PutValue(row * 10);   // some numeric value
            }

            // Add a chart on the source sheet that uses the original data range A2:B20
            int chartIndex = srcSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = srcSheet.Charts[chartIndex];
            // Set the initial data range (source sheet)
            chart.SetChartDataRange($"{srcSheet.Name}!A2:B20", true);

            // Add a destination worksheet where rows will be copied
            Worksheet destSheet = workbook.Worksheets.Add("Destination");

            // Copy rows 2‑20 from source sheet to destination sheet starting at row 1 (Excel row 1)
            // sourceRowIndex = 1 (row 2), destinationRowIndex = 0 (row 1), rowNumber = 19 rows
            destSheet.Cells.CopyRows(srcSheet.Cells, 1, 0, 19);

            // Update the chart to refer to the new data location on the destination sheet
            // The copied data occupies rows 1‑19 on the destination sheet (A1:B19)
            // To keep the same relative range (A2:B20) we offset by one row
            chart.SetChartDataRange($"{destSheet.Name}!A2:B20", true);

            // Optionally move the chart to the destination sheet for visual consistency
            // (charts are objects attached to a worksheet; moving requires recreating or copying)
            // Here we simply reposition it within the source sheet
            chart.Move(20, 0, 30, 5);

            // Save the workbook
            workbook.Save("CopyRowsAndUpdatedChart.xlsx");
        }
    }
}