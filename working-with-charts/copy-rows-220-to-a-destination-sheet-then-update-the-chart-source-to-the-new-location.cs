// Title: Copy rows 2‑20 from a source worksheet to another sheet and rebind a column chart to the new range using Aspose.Cells for .NET (C#)
// AI Prompts: Copy rows 2‑20 from the 'Source' worksheet to the 'Destination' worksheet and then set the chart's data range to Destination!A1:B19 with Aspose.Cells in C#. | After moving a data block to a new sheet, update an existing column chart to reference the copied range using the SetChartDataRange method in Aspose.Cells. | Create a workbook that populates rows, adds a column chart, copies those rows to a different worksheet, and reassigns the chart source programmatically in C#.
// Common Searches: Aspose.Cells C# copy rows from one worksheet to another and change chart source range | how to update chart data range after copying rows with Aspose.Cells | SetChartDataRange to Destination sheet after copying rows in Aspose.Cells .NET | copy rows 2 to 20 and rebind column chart in Aspose.Cells C# example
// Tags: copy rows between worksheets Aspose.Cells C# | update chart data range Aspose.Cells | SetChartDataRange method example | column chart source adjustment Aspose.Cells | copy rows and rebind chart Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCopyRowsAndUpdateChart
{
    // The program creates a workbook, fills rows 2‑20 on a 'Source' sheet, adds a column chart referencing that range, copies the rows to a new 'Destination' sheet, updates the chart to point to Destination!A1:B19, and saves the file as CopyRowsAndUpdateChart.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (source)
            Workbook workbook = new Workbook();
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "Source";

            // Populate rows 2‑20 (indices 1‑19) with sample data in columns A and B
            for (int row = 1; row <= 19; row++) // row index 1 = Excel row 2
            {
                sourceSheet.Cells[row, 0].PutValue($"Item {row}");
                sourceSheet.Cells[row, 1].PutValue(row * 10);
            }

            // Add a chart that uses the data from rows 2‑20 on the source sheet
            int chartIndex = sourceSheet.Charts.Add(ChartType.Column, 22, 0, 32, 5);
            Chart chart = sourceSheet.Charts[chartIndex];
            // Initial data range points to the source sheet
            chart.SetChartDataRange("Source!A2:B20", true);

            // Add a destination worksheet
            Worksheet destSheet = workbook.Worksheets.Add("Destination");

            // Copy rows 2‑20 from source to destination starting at row 1 (Excel row 1)
            // Source row index = 1 (row 2), destination row index = 0 (row 1), number of rows = 19
            destSheet.Cells.CopyRows(sourceSheet.Cells, 1, 0, 19);

            // Update the chart's data source to refer to the copied range on the destination sheet
            // The copied range occupies rows 1‑19 (Excel rows 1‑19) in the destination sheet
            // Adjust the range to match the actual rows that contain data (rows 1‑19)
            chart.SetChartDataRange("Destination!A1:B19", true);

            // Save the workbook
            workbook.Save("CopyRowsAndUpdateChart.xlsx");
        }
    }
}
