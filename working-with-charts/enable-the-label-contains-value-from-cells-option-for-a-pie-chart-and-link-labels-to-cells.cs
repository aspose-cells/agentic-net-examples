// Title: Enable the "Label Contains – Value From Cells" option for a pie chart and link its data labels to worksheet cells using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a pie chart with Aspose.Cells and sets the DataLabels.LinkedSource property to a cell range so the labels display custom text from the worksheet. | Show how to activate the "Label Contains – Value From Cells" feature for a pie chart series and configure the chart to avoid overlapping data labels in Aspose.Cells. | Adapt the example to pull label text from a different column or range while keeping the linked‑source behavior for the chart series.
// Common Searches: Aspose.Cells C# link pie chart data labels to cells | Enable label contains value from cells option for Excel chart using Aspose.Cells | Set DataLabels.LinkedSource property in Aspose.Cells pie chart | Prevent overlapping data labels in Aspose.Cells pie chart | Use custom label column for chart series Aspose.Cells .NET
// Tags: pie chart data labels linked source Aspose.Cells | enable label contains value from cells Aspose.Cells | prevent overlapping labels pie chart Aspose.Cells | custom label column for chart series Aspose.Cells | Aspose.Cells C# create pie chart from worksheet data

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsPieChartLabelFromCells
{
    // Demonstrates creating a workbook, adding sample data, inserting a pie chart, and linking its data labels to custom text cells (C2:C4) by enabling the "Label Contains – Value From Cells" option with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Column A – Category names
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Cherry");

            // Column B – Numeric values for the pie slices
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["B4"].PutValue(100);

            // Column C – Custom label text that we want the data labels to display
            sheet.Cells["C1"].PutValue("Label");
            sheet.Cells["C2"].PutValue("Apple – 120 units");
            sheet.Cells["C3"].PutValue("Banana – 80 units");
            sheet.Cells["C4"].PutValue("Cherry – 100 units");

            // Add a pie chart to the worksheet
            int chartIdx = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 10);
            Chart pieChart = sheet.Charts[chartIdx];

            // Add the series (values) and associate category names
            Series series = pieChart.NSeries[pieChart.NSeries.Add("=Sheet1!$B$2:$B$4", true)];
            series.XValues = "=Sheet1!$A$2:$A$4";

            // Enable data labels and link them to the custom label cells (C2:C4)
            series.DataLabels.ShowValue = true;                 // Show the value (optional when using LinkedSource)
            series.DataLabels.LinkedSource = "C2:C4";            // Link label text to cells
            series.DataLabels.NumberFormatLinked = true;        // Keep number format in sync with source cells
            series.DataLabels.IsNeverOverlap = true;            // Prevent overlapping labels in a pie chart

            // Save the workbook
            workbook.Save("PieChart_LabelFromCells.xlsx");
        }
    }
}
