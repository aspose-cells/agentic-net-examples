// Title: How to bind custom data label text from a worksheet range to the first series of a column chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that sets the first series of a column chart to use cell range C2:C4 as the source for its data label text. | Show how to enable data labels, link them to a worksheet range, and style them (position outside end, blue font) for a chart series in Aspose.Cells. | Create a workbook, add sample data, generate a column chart, and configure the first series to display custom labels from cells using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# bind chart series data labels to a cell range | set custom label text for column chart series from worksheet cells in .NET | how to use LinkedSource property for chart data labels with Aspose.Cells | display units in chart data labels using Aspose.Cells C# example | Aspose.Cells chart series label position outside end code sample
// Tags: chart series data labels from worksheet range Aspose.Cells | column chart custom label text C# | configure data label position outside end Aspose.Cells | save workbook with chart as xlsx Aspose.Cells | set data labels show value Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsDataLabelsDemo
{
    // The example creates a new workbook, adds category, value, and custom label columns, inserts a column chart, and configures the first series to show data labels whose text is taken from the cell range C2:C4. Labels are positioned outside the end of each column and styled with a blue font. The workbook is saved as DataLabelsFromCellRange.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["B4"].PutValue(300);

            // Cells that contain custom label texts (e.g., with units)
            sheet.Cells["C1"].PutValue("Label");
            sheet.Cells["C2"].PutValue("100 units");
            sheet.Cells["C3"].PutValue("200 units");
            sheet.Cells["C4"].PutValue("300 units");

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (values) and categories
            chart.NSeries.Add("=Sheet1!$B$2:$B$4", true);
            chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$4";

            // Access the first series
            Series firstSeries = chart.NSeries[0];

            // Enable data labels and configure them to use a cell range as the source
            firstSeries.DataLabels.ShowValue = true;          // Show the numeric value (optional)
            firstSeries.DataLabels.ShowCellRange = true;      // Use cell range for label text
            firstSeries.DataLabels.LinkedSource = "C2:C4";    // Range containing custom label texts
            firstSeries.DataLabels.Position = LabelPositionType.OutsideEnd;
            firstSeries.DataLabels.Font.Color = Color.Blue;

            // Save the workbook
            workbook.Save("DataLabelsFromCellRange.xlsx", SaveFormat.Xlsx);
        }
    }
}
