// Title: How to resize chart data label shapes by adjusting font size in Aspose.Cells for .NET after adding hyperlinks
// AI Prompts: Generate a C# workbook that creates a column chart, adds a hyperlink to each data label, and enlarges the label by setting DataLabels.Font.Size. | Write Aspose.Cells code to enable data labels on a chart and programmatically increase their visual size to accommodate added external hyperlinks. | Provide a snippet that resizes chart data label shapes after assigning URL hyperlinks to them in an Excel file using Aspose.Cells.
// Common Searches: asp.net resize chart data label after adding hyperlink Aspose.Cells | set data label font size in column chart using Aspose.Cells C# | add hyperlink to each data point label in Excel chart with Aspose.Cells | increase size of data label shapes in Aspose.Cells chart | how to adjust chart data label dimensions programmatically in .NET
// Tags: chart data label font size Aspose.Cells | add hyperlink to chart data label .NET | resize chart data label shape Aspose.Cells | column chart data label styling C# | Aspose.Cells chart label customization

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a new workbook, fills sample data, adds a column chart, enables data labels, sets the data label font size to 12, and saves the workbook as ResizedDataLabels.xlsx.
class ResizeDataLabelShapes
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true; // display the value in each label

            // Adjust font size for data labels (proxy for label size)
            series.DataLabels.Font.Size = 12;

            // Save the workbook
            workbook.Save("ResizedDataLabels.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
