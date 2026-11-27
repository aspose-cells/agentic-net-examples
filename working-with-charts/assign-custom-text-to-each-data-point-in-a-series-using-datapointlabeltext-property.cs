// Title: How to assign custom text to individual chart data points using DataLabels.Text in Aspose.Cells for .NET (C#)
// Common Searches: C# set custom label text for each point in an Aspose.Cells column chart | Aspose.Cells DataLabels.Text example for displaying category and value on chart points | programmatically change individual data point labels in an Aspose.Cells chart | how to position chart data labels inside the column end using Aspose.Cells | customize font color and style of data point labels in Aspose.Cells C#
// Tags: DataLabels.Text custom label Aspose.Cells | chart point label customization C# | Aspose.Cells column chart data point labels | label font color and bold Aspose.Cells | inside end label position Aspose.Cells | save workbook as XLSX Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsDataPointLabelDemo
{
    // The sample creates a workbook, fills columns with categories and values, adds a column chart, disables the default value and category labels, then loops through each ChartPoint to set DataLabels.Text to a string that combines the category name and its numeric value, applies dark‑blue bold font, positions the label inside the column end, enables text wrapping and transparent background, and finally saves the file as an XLSX workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the chart
            // Column A – Category names
            // Column B – Values
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");
            cells["A2"].PutValue("Alpha");
            cells["A3"].PutValue("Beta");
            cells["A4"].PutValue("Gamma");
            cells["A5"].PutValue("Delta");
            cells["B2"].PutValue(40);
            cells["B3"].PutValue(70);
            cells["B4"].PutValue(55);
            cells["B5"].PutValue(90);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the series
            // Use the values column as Y values and the category column as X values
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Ensure that data labels are visible (optional, helps to see the custom text)
            chart.NSeries[0].DataLabels.ShowValue = false; // we will show custom text instead
            chart.NSeries[0].DataLabels.ShowCategoryName = false;

            // Iterate through each data point in the first series
            // and assign a custom label text
            ChartPointCollection points = chart.NSeries[0].Points;
            for (int i = 0; i < points.Count; i++)
            {
                ChartPoint point = points[i];

                // Access the DataLabels object for the point
                DataLabels pointLabel = point.DataLabels;

                // Set custom text – you can build any string you need
                pointLabel.Text = $"Item {i + 1}: {cells[$"A{i + 2}"].StringValue} = {cells[$"B{i + 2}"].IntValue}";

                // Optionally adjust appearance of the label
                pointLabel.Font.Color = Color.DarkBlue;
                pointLabel.Font.IsBold = true;
                pointLabel.Position = LabelPositionType.InsideEnd; // place label inside the column
                pointLabel.IsTextWrapped = true;
                pointLabel.BackgroundMode = BackgroundMode.Transparent;
            }

            // Save the workbook to an XLSX file
            string outputPath = "DataPointCustomLabels.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
            Console.WriteLine("Each data point now displays a custom label defined via DataPoint.Label.Text (DataLabels.Text).");
        }
    }
}
