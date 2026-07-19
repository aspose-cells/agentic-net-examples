// Title: Aspose.Cells C# – Set Column Chart Data Labels to OutsideEnd and Add Leader Lines
// Description: Creates a workbook, adds a column chart with categories and values, enables data labels for the first series, positions them at OutsideEnd to keep them clear of the bars, activates leader lines, and saves the file as DataLabelOutsideEndDemo.xlsx.
// Keywords: Aspose.Cells | C# chart data labels | OutsideEnd label position | leader lines | column chart | Excel automation | chart label positioning | Aspose.Cells API | data label overlap | chart customization
// Common Searches: Aspose.Cells set data label position OutsideEnd C# | Enable leader lines in Aspose.Cells chart | C# column chart label outside end Aspose | Prevent data label overlap Aspose.Cells | How to move chart data labels outside bars Aspose.Cells
// Developer Intent: Configure a chart so that data labels appear outside the columns with leader lines, eliminating label‑to‑bar overlap.
// Use Cases: Financial report showing quarterly sales where each column’s value is displayed outside the bar for readability. | Automated dashboard that generates Excel files with multi‑series charts and clear label placement. | Exporting data to Excel for presentations, ensuring labels do not obscure chart elements.
// AI Prompts: Generate C# code using Aspose.Cells to create a column chart and set data labels to OutsideEnd with leader lines. | Show how to apply the OutsideEnd label position to all series in an Aspose.Cells chart and toggle leader line visibility. | Explain step‑by‑step how to avoid data label overlap in Aspose.Cells charts by using OutsideEnd positioning and leader lines.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsDataLabelPositionDemo
{
    // Creates a workbook, adds a column chart with categories and values, enables data labels for the first series, positions them at OutsideEnd to keep them clear of the bars, activates leader lines, and saves the file as DataLabelOutsideEndDemo.xlsx.
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

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;

            // Position data labels outside the end of the bars to avoid overlapping leader lines
            series.DataLabels.Position = LabelPositionType.OutsideEnd;

            // Enable leader lines (optional, but often used with OutsideEnd)
            series.HasLeaderLines = true;

            // Save the workbook
            workbook.Save("DataLabelOutsideEndDemo.xlsx");
        }
    }
}
