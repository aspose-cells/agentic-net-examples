// Title: Assign custom rich‑text data labels to each point in a column chart using Aspose.Cells for C#
// Common Searches: Aspose.Cells C# set custom rich text label for each chart point | how to format individual data labels in a column chart with Aspose.Cells | C# example of applying bold red text to part of a chart data label in Aspose.Cells | hide default values and show custom labels on Aspose.Cells column chart points | change font color and style of chart point labels using Aspose.Cells API
// Tags: custom rich‑text chart data labels Aspose.Cells | format chart point label characters C# | hide default data label values Aspose.Cells | column chart series point annotation Aspose.Cells | apply font settings to chart data labels C#

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AssignRichTextLabels
{
    // The sample creates a workbook, fills category and numeric data, adds a column chart, and calculates it so the points collection is available. It then iterates through each point, builds a custom label like "Point 1: Item 1", hides the default value and category, and applies rich‑text formatting—bold red for the word "Point", blue italic for the index, and dark‑green for the category name. Finally, the workbook is saved as RichTextLabelsChart.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // 1. Populate sample data (Category / Value)
            // ------------------------------------------------------------
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            for (int i = 2; i <= 6; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                sheet.Cells[$"B{i}"].PutValue(10 * i); // arbitrary numeric values
            }

            // ------------------------------------------------------------
            // 2. Add a column chart to the worksheet
            // ------------------------------------------------------------
            // Parameters: chart type, upper left row, upper left column, lower right row, lower right column
            int chartIndex = sheet.Charts.Add(ChartType.Column, 8, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the series
            // Values are in B2:B6, categories are in A2:A6
            chart.NSeries.Add("B2:B6", true);
            chart.NSeries.CategoryData = "A2:A6";

            // ------------------------------------------------------------
            // 3. Assign a unique rich‑text label to each data point
            // ------------------------------------------------------------
            // Ensure the chart is calculated so that points collection is populated
            chart.Calculate();

            // Access the first (and only) series
            Series series = chart.NSeries[0];

            // Iterate through each point in the series
            for (int i = 0; i < series.Points.Count; i++)
            {
                ChartPoint point = series.Points[i];

                // ----------------------------------------------------------------
                // Create a custom label for this point
                // ----------------------------------------------------------------
                string labelText = $"Point {i + 1}: {sheet.Cells[$"A{i + 2}"].StringValue}";
                point.DataLabels.ShowValue = false;          // hide the default numeric value
                point.DataLabels.ShowCategoryName = false;   // hide the default category name
                point.DataLabels.Text = labelText;           // set our custom label

                // ----------------------------------------------------------------
                // Apply rich‑text formatting to parts of the label
                // ----------------------------------------------------------------
                // Example: make the word "Point" bold and red, the rest normal
                // Characters(startIndex, length) returns a FontSetting that can be formatted
                FontSetting pointWord = point.DataLabels.Characters(0, 5); // "Point"
                pointWord.Font.IsBold = true;
                pointWord.Font.Color = Color.Red;
                pointWord.Font.Size = 12;

                // Example: make the numeric index (e.g., "1") blue and italic
                int indexStart = labelText.IndexOf((i + 1).ToString());
                if (indexStart >= 0)
                {
                    FontSetting indexWord = point.DataLabels.Characters(indexStart, (i + 1).ToString().Length);
                    indexWord.Font.Color = Color.Blue;
                    indexWord.Font.IsItalic = true;
                }

                // Example: set the category name part (after colon) to dark green
                int colonPos = labelText.IndexOf(':');
                if (colonPos >= 0 && colonPos + 2 < labelText.Length)
                {
                    FontSetting categoryWord = point.DataLabels.Characters(colonPos + 2, labelText.Length - colonPos - 2);
                    categoryWord.Font.Color = Color.DarkGreen;
                }
            }

            // ------------------------------------------------------------
            // 4. Save the workbook
            // ------------------------------------------------------------
            string outputPath = "RichTextLabelsChart.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
