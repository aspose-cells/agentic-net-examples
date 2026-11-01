// Title: Rotate primary Y‑axis tick labels 90° clockwise in an Aspose.Cells column chart using C#
// AI Prompts: Generate C# code that accesses Chart.ValueAxis.TickLabels and sets RotationAngle to 90 for a column chart created with Aspose.Cells. | Show how to retrieve the primary Y‑axis TickLabels object of an Aspose.Cells chart and rotate the labels clockwise. | Provide an example that changes the orientation of Y‑axis tick labels to vertical in a .NET Aspose.Cells workbook.
// Common Searches: Aspose.Cells C# rotate primary Y axis tick labels 90 degrees | set chart value axis label rotation angle Aspose.Cells .NET | vertical Y‑axis tick labels for column chart using Aspose.Cells | how to change tick label orientation in an Aspose.Cells Excel chart
// Tags: valueaxis ticklabels rotation aspnet | aspocells chart yaxis label orientation | c# column chart label angle aspocells | excel chart tick label rotation aspocells | primary y axis formatting aspocells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a column chart, accesses the primary Y‑axis (ValueAxis) TickLabels object, sets its RotationAngle to 90 degrees clockwise, and saves the file as TickLabelsRotated.xlsx.
class Program
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
            sheet.Cells["A2"].PutValue("Item 1");
            sheet.Cells["A3"].PutValue("Item 2");
            sheet.Cells["A4"].PutValue("Item 3");
            sheet.Cells["A5"].PutValue("Item 4");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(40);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart series
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Rotate Y‑axis tick labels 90 degrees clockwise
            chart.ValueAxis.TickLabels.RotationAngle = 90;

            // Save the workbook
            workbook.Save("TickLabelsRotated.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
