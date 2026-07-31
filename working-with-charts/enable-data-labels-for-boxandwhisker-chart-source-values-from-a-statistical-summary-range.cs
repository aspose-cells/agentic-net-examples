// Title: Add Centered Data Labels to a Box‑Whisker Chart from a Summary Range with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills category and summary columns, inserts a Box‑Whisker chart, uses SetChartDataRange with the summary flag, assigns category data, enables ShowValue on the first series, links the labels to the same summary range, positions them at the center, and saves the file as BoxWhiskerDataLabels.xlsx.
// Keywords: Aspose.Cells C# box whisker chart | data labels linked source | summary range chart Aspose.Cells | .NET chart data labels | BoxWhisker chart label positioning | Aspose.Cells tutorial | Excel chart automation | US developers Aspose.Cells
// Common Searches: Aspose.Cells show data labels on box whisker chart | C# set linked source for chart data labels Aspose.Cells | centered data labels box whisker Aspose.Cells .NET | use summary range for Aspose.Cells chart data
// Developer Intent: Enable data labels on a Box‑Whisker chart and bind them to a statistical summary range using Aspose.Cells for .NET.
// Use Cases: Highlight median values directly on quarterly performance box‑whisker charts. | Display custom statistics (min, max, average) as labels for each category in financial reports. | Provide a clean visual summary by linking labels to a separate summary column in multi‑category analyses.
// AI Prompts: Write C# code with Aspose.Cells that adds a Box‑Whisker chart, sets a summary data range, and shows centered data labels linked to that range. | Show how to configure ShowValue and LinkedSource for a chart series in Aspose.Cells. | Explain the purpose of the summary flag in SetChartDataRange and how it affects data labels.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsBoxWhiskerDataLabels
{
    // Creates a workbook, fills category and summary columns, inserts a Box‑Whisker chart, uses SetChartDataRange with the summary flag, assigns category data, enables ShowValue on the first series, links the labels to the same summary range, positions them at the center, and saves the file as BoxWhiskerDataLabels.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Category column
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q1");
            sheet.Cells["A4"].PutValue("Q1");
            sheet.Cells["A5"].PutValue("Q2");
            sheet.Cells["A6"].PutValue("Q2");
            sheet.Cells["A7"].PutValue("Q2");

            // Values column (statistical summary for each category)
            sheet.Cells["B1"].PutValue("Values");
            sheet.Cells["B2"].PutValue(15);
            sheet.Cells["B3"].PutValue(25);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(20);
            sheet.Cells["B6"].PutValue(35);
            sheet.Cells["B7"].PutValue(40);

            // Add a Box‑Whisker chart
            int chartIndex = sheet.Charts.Add(ChartType.BoxWhisker, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Use the statistical summary range as the data source (true = summary data)
            chart.SetChartDataRange("B2:B7", true);
            chart.NSeries.CategoryData = "A2:A7";

            // Enable data labels for the first series and link them to the same summary range
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;               // display the values
            series.DataLabels.Position = LabelPositionType.Center; // optional positioning
            series.DataLabels.LinkedSource = "B2:B7";         // source values for labels

            // Save the workbook
            workbook.Save("BoxWhiskerDataLabels.xlsx");
        }
    }
}
