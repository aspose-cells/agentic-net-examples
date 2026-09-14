// Title: How to assign unique rich‑text data labels to each point of a column chart using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a column chart in an Aspose.Cells workbook and sets a multi‑line rich‑text label for every chart point. | Provide a C# snippet that changes the font color and applies bold styling to individual data labels in an Aspose.Cells chart. | Show how to adapt the example to a line chart and include both the category name and series name in each custom label with Aspose.Cells.
// Common Searches: Aspose.Cells C# set custom rich text label for each data point in a column chart | how to change font color of individual data labels in Aspose.Cells chart using .NET | assign different text to each point of a chart series with Aspose.Cells for .NET | disable automatic data label text and use custom labels in Aspose.Cells charts
// Tags: rich‑text data labels Aspose.Cells | per‑point chart label styling Aspose.Cells | column chart label font color Aspose.Cells | suppress default data label text Aspose.Cells | save workbook as XLSX Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsRichTextDataLabels
{
    // The example creates a new workbook, populates it with category and value data, adds a column chart, binds the data series, disables automatic label text, and then iterates over each chart point to assign a custom multi‑line rich‑text label that includes the item number, category, and value. It alternates the label font color between blue and green and makes the text bold before saving the file as RichTextDataLabels.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (Category and Value)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Alpha");
            sheet.Cells["A3"].PutValue("Beta");
            sheet.Cells["A4"].PutValue("Gamma");
            sheet.Cells["A5"].PutValue("Delta");
            sheet.Cells["B2"].PutValue(15);
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["B4"].PutValue(45);
            sheet.Cells["B5"].PutValue(60);

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIdx];

            // Bind data to the chart
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;          // Show the numeric value
            series.DataLabels.ShowCategoryName = false; // We'll replace with custom text
            series.DataLabels.IsAutoText = false;       // Allow custom text per point

            // Iterate through each data point and assign a unique rich‑text label
            for (int i = 0; i < series.Points.Count; i++)
            {
                ChartPoint point = series.Points[i];

                // Build a custom label (you can embed any rich‑text formatting you need)
                string customLabel = $"Item {i + 1}\nCategory: {point.XValue}\nValue: {point.YValue}";

                // Assign the custom text to the point's data label
                point.DataLabels.Text = customLabel;

                // Optional: customize font for this label (e.g., different color per point)
                point.DataLabels.Font.Color = i % 2 == 0 ? Color.Blue : Color.Green;
                point.DataLabels.Font.IsBold = true;
            }

            // Save the workbook (XLSX format)
            workbook.Save("RichTextDataLabels.xlsx");
        }
    }
}
