// Title: C# – Assign Custom Text to Each Chart Point in Aspose.Cells
// Description: Demonstrates how to create a workbook, add a column chart, enable data labels, and replace the default auto‑generated label with a custom string for every ChartPoint using the DataLabels.Text property, then save the file as XLSX.
// Keywords: Aspose.Cells custom data labels | C# chart point label text | DataLabels.Text Aspose.Cells | disable auto text chart labels .NET | set individual point labels Aspose.Cells
// Common Searches: Aspose.Cells set custom label for each chart point | C# change data label text in column chart Aspose | how to disable auto text in Aspose.Cells chart | custom data labels per series Aspose.Cells .NET
// Developer Intent: Replace the default data‑label text with a unique string for every point in a chart series.
// Use Cases: Show descriptive labels such as "Point 1: 10" instead of raw values for clearer reporting. | Add prefixes, units, or indexes to chart labels to match corporate style guides. | Create numbered annotations on a chart when the underlying data lacks meaningful names.
// AI Prompts: Generate C# code that iterates over a chart series in Aspose.Cells and sets a custom DataLabels.Text for each point. | Explain how to turn off auto‑generated labels and assign formatted strings to ChartPoint.DataLabels in a column chart. | Provide a step‑by‑step example of customizing data label text for every data point using Aspose.Cells .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCustomDataLabels
{
    // Demonstrates how to create a workbook, add a column chart, enable data labels, and replace the default auto‑generated label with a custom string for every ChartPoint using the DataLabels.Text property, then save the file as XLSX.
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
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true; // show the default value

            // Assign custom text to each data point's label
            for (int i = 0; i < series.Points.Count; i++)
            {
                ChartPoint point = series.Points[i];
                // Disable auto-generated text and set a custom label
                point.DataLabels.IsAutoText = false;
                point.DataLabels.Text = $"Point {i + 1}: {point.YValue}";
            }

            // Save the workbook to an XLSX file
            workbook.Save("CustomDataLabels.xlsx");
        }
    }
}
