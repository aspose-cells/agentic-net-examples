using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDataLabelsFormatting
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the first series and enable data labels
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;

            // Apply bold font style to all data labels
            series.DataLabels.Font.IsBold = true;

            // Center align the text inside data labels (both horizontally and vertically)
            series.DataLabels.TextHorizontalAlignment = TextAlignmentType.Center;
            series.DataLabels.TextVerticalAlignment = TextAlignmentType.Center;

            // Apply the font settings to all child data label nodes
            series.DataLabels.ApplyFont();

            // Save the workbook
            workbook.Save("DataLabelsBoldCentered.xlsx");
        }
    }
}