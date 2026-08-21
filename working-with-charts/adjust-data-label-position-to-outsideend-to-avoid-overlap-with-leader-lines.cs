// Title: Aspose.Cells C# – Position Column Chart Data Labels OutsideEnd to Prevent Overlap
// Description: Demonstrates how to create a workbook, add a column chart, enable data labels, set their position to OutsideEnd, optionally show leader lines, and save the file as DataLabelPositionOutsideEnd.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells data label position | LabelPositionType OutsideEnd | column chart data labels C# | prevent label overlap Aspose | leader lines Aspose.Cells | chart formatting Aspose.Cells | C# Excel chart example
// Common Searches: Aspose.Cells set data label position OutsideEnd | C# column chart label outside end | avoid data label overlap Aspose.Cells | enable leader lines in Aspose.Cells chart | change chart data label placement .NET
// Developer Intent: Place column chart data labels outside the column ends to keep them clear of leader lines.
// Use Cases: Generate a column chart from worksheet data and apply series.DataLabels.Position = LabelPositionType.OutsideEnd for better readability. | Activate series.HasLeaderLines = true when visual connectors between columns and labels are required. | Automate Excel report creation with properly positioned data labels to improve visual clarity.
// AI Prompts: Write C# code with Aspose.Cells that creates a column chart and sets data label position to OutsideEnd while enabling leader lines. | Explain the effect of LabelPositionType.OutsideEnd on data label placement in column charts and how to toggle leader lines in Aspose.Cells. | Provide a step‑by‑step guide to adjust data label positions for multiple series in an Aspose.Cells chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsDataLabelPositionDemo
{
    // Demonstrates how to create a workbook, add a column chart, enable data labels, set their position to OutsideEnd, optionally show leader lines, and save the file as DataLabelPositionOutsideEnd.xlsx using Aspose.Cells for .NET.
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

            // Position data labels outside the end of each column to avoid overlapping leader lines
            series.DataLabels.Position = LabelPositionType.OutsideEnd;

            // Optional: enable leader lines if needed
            series.HasLeaderLines = true;

            // Save the workbook to an Excel file
            workbook.Save("DataLabelPositionOutsideEnd.xlsx");
        }
    }
}
