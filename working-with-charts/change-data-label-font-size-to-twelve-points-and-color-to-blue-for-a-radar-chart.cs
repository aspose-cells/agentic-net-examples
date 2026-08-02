// Title: C# – Set Radar Chart Data Labels to 12‑pt Blue Font with Aspose.Cells
// Description: Creates a workbook, adds sample data, inserts a Radar chart, enables data labels for the first series, and formats those labels with a 12‑point blue font before saving the file as an XLSX document.
// Keywords: Aspose.Cells | C# | .NET | Radar chart | Data labels | Font size 12 | Blue font color | Chart customization | Excel export | ChartType.Radar | Series.DataLabels | ApplyFont
// Common Searches: Aspose.Cells set radar chart data label font size | C# change radar chart label color to blue | How to format data labels in Aspose.Cells chart | Apply 12 pt font to chart series labels .NET | Radar chart label styling example Aspose.Cells
// Developer Intent: Apply a 12‑point blue font to all data labels of a radar chart series using Aspose.Cells for .NET.
// Use Cases: Generate Excel reports with radar charts that match corporate branding by standardizing label appearance. | Automate creation of dashboards where radar chart data labels need clear, readable styling. | Provide a reusable code snippet for developers who need consistent label formatting across multiple charts.
// AI Prompts: Write C# code with Aspose.Cells that creates a radar chart and sets its data label font to 12 points in blue. | Explain step‑by‑step how to enable and style data labels for a radar chart series in Aspose.Cells for .NET. | Create a helper method that receives a Chart object and applies a 12‑pt blue font to all its data labels.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, inserts a Radar chart, enables data labels for the first series, and formats those labels with a 12‑point blue font before saving the file as an XLSX document.
class RadarChartDataLabelFormatting
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the radar chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Cat1");
        sheet.Cells["A3"].PutValue("Cat2");
        sheet.Cells["A4"].PutValue("Cat3");
        sheet.Cells["A5"].PutValue("Cat4");

        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["B2"].PutValue(4);
        sheet.Cells["B3"].PutValue(2);
        sheet.Cells["B4"].PutValue(5);
        sheet.Cells["B5"].PutValue(3);

        // Add a radar chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Radar, 6, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data source for the chart
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Enable data labels for the first series
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;

        // Change data label font size to 12 points and color to blue
        series.DataLabels.Font.Size = 12;
        series.DataLabels.Font.Color = Color.Blue;

        // Apply the font settings to all data labels
        series.DataLabels.ApplyFont();

        // Save the workbook
        workbook.Save("RadarChartDataLabels.xlsx", SaveFormat.Xlsx);
    }
}
