// Title: Add custom text labels from a worksheet column to a column chart in Aspose.Cells for .NET (C#)
// AI Prompts: Create a column chart and assign each point's data label to the string value stored in column C of the worksheet. | Configure the chart to hide the default category names and display only the custom labels taken from a specified cell range. | Save the workbook as an .xlsx file after applying the custom data labels to the chart series.
// Common Searches: Aspose.Cells C# set chart point labels from another column | How to replace category names with custom text in an Aspose.Cells column chart | C# Aspose.Cells chart hide category axis labels and use cell values as data labels | Assign custom data labels to chart series points using Aspose.Cells for .NET
// Tags: custom data labels from cells Aspose.Cells | column chart label range C# Aspose.Cells | hide default category names chart Aspose.Cells | assign point label worksheet column Aspose.Cells | Aspose.Cells chart series custom labels

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The example creates a workbook, fills columns with categories, values, and custom label text, adds a column chart, hides the default category names, assigns each chart point a custom data label from column C, and saves the result as ChartWithCustomLabels.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Header row
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["C1"].PutValue("Label");

            // Sample data
            string[] categories = { "Jan", "Feb", "Mar", "Apr" };
            double[] values = { 10, 20, 15, 30 };
            string[] labels = { "January", "February", "March", "April" };

            // Populate worksheet
            for (int i = 0; i < categories.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(categories[i]); // Column A
                sheet.Cells[i + 1, 1].PutValue(values[i]);    // Column B
                sheet.Cells[i + 1, 2].PutValue(labels[i]);   // Column C
            }

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set series values (B2:B5). Categories will be taken from the first column by default.
            chart.NSeries.Add("B2:B5", true);

            // Show values as data labels, hide default category names
            chart.NSeries[0].DataLabels.ShowValue = true;
            chart.NSeries[0].DataLabels.ShowCategoryName = false;

            // Apply custom labels from column C to each point
            int pointCount = chart.NSeries[0].Points.Count;
            for (int i = 0; i < pointCount; i++)
            {
                string customLabel = sheet.Cells[i + 1, 2].StringValue;
                // Set custom data label for each point
                chart.NSeries[0].Points[i].DataLabels.Text = customLabel;
            }

            // Save the workbook
            workbook.Save("ChartWithCustomLabels.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
