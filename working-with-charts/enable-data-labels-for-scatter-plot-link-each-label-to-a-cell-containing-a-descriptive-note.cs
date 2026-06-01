using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsScatterDataLabels
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate X values, Y values and descriptive notes
            // X values in column A, Y values in column B, notes in column C
            sheet.Cells["A1"].PutValue("X");
            sheet.Cells["B1"].PutValue("Y");
            sheet.Cells["C1"].PutValue("Note");

            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["C2"].PutValue("First point");

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["C3"].PutValue("Second point");

            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue(15);
            sheet.Cells["C4"].PutValue("Third point");

            sheet.Cells["A5"].PutValue(4);
            sheet.Cells["B5"].PutValue(25);
            sheet.Cells["C5"].PutValue("Fourth point");

            // Add a scatter chart
            int chartIndex = sheet.Charts.Add(ChartType.Scatter, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set X and Y data for the series
            chart.NSeries.Add("B2:B5", true);          // Y values
            chart.NSeries[0].XValues = "A2:A5";       // X values

            // Enable data labels and link each label to the corresponding note cell
            Series series = chart.NSeries[0];
            series.DataLabels.ShowCellRange = true;    // Use cell range for label text
            series.DataLabels.LinkedSource = "C2:C5"; // Range containing notes
            series.DataLabels.Position = LabelPositionType.Above; // Optional positioning
            series.DataLabels.Font.Color = Color.DarkBlue;       // Optional styling

            // Save the workbook
            workbook.Save("ScatterChartWithLinkedDataLabels.xlsx");
        }
    }
}