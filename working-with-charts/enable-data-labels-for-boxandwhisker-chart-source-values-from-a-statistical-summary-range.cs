// Title: Aspose.Cells for .NET – Add Data Labels to a Box‑and‑Whisker Chart Using a Statistical Summary Range
// Description: Shows how to build a workbook, fill category and raw data, insert a Box‑Whisker chart in statistical‑summary mode, enable data labels that display the calculated values, set the label position, and save the file as XLSX with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# box whisker chart | data labels | statistical summary range | show values | label position | BoxWhisker chart .NET | Excel chart automation | chart data labels Aspose | box plot data labels
// Common Searches: Aspose.Cells enable data labels on box plot | C# box whisker chart statistical summary Aspose | How to show values in Aspose.Cells box‑whisker chart | Set label position for box plot in Aspose.Cells | Create box‑whisker chart from summary range using Aspose.Cells
// Developer Intent: Create a Box‑Whisker chart from a statistical summary range and display the computed statistics as data labels, with optional positioning.
// Use Cases: Financial reporting: visualize quarterly performance with a box‑whisker chart that shows median and quartile values as centered labels. | Quality‑control dashboard: plot summary statistics for production batches and label each box with its key values for quick inspection. | Scientific data presentation: export experimental results to Excel where each box‑whisker plot includes data labels that highlight calculated metrics.
// AI Prompts: Generate C# code with Aspose.Cells that adds a Box‑Whisker chart from a statistical summary range and enables data labels showing the calculated values at the center. | Provide an Aspose.Cells example that customizes the font size, color, and background of data labels for a Box‑Whisker chart. | Explain how to configure individual data label settings for multiple series in a Box‑Whisker chart using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to build a workbook, fill category and raw data, insert a Box‑Whisker chart in statistical‑summary mode, enable data labels that display the calculated values, set the label position, and save the file as XLSX with Aspose.Cells for .NET.
class BoxWhiskerDataLabelsDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for a box‑and‑whisker chart
        // Column A – categories (e.g., quarters)
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q1");
        sheet.Cells["A4"].PutValue("Q1");
        sheet.Cells["A5"].PutValue("Q2");
        sheet.Cells["A6"].PutValue("Q2");
        sheet.Cells["A7"].PutValue("Q2");

        // Column B – raw values that will be treated as a statistical summary
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

        // Use the statistical summary range (the second parameter = true)
        chart.SetChartDataRange("B2:B7", true);
        chart.NSeries.CategoryData = "A2:A7";

        // Enable data labels for the first series and show the calculated values
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;
        // Optional: set the position of the data labels
        series.DataLabels.Position = LabelPositionType.Center;

        // Save the workbook
        workbook.Save("BoxWhiskerDataLabels.xlsx", SaveFormat.Xlsx);
    }
}
