// Title: Aspose.Cells .NET: Set and Read a Chart Subtitle
// Description: Creates a workbook, adds sample data, inserts a column chart, assigns a main title and a custom subtitle, reads back the subtitle to verify the change, prints it, and saves the file as ChartWithSubtitle.xlsx using Aspose.Cells for C#.
// Keywords: Aspose.Cells chart subtitle | C# set chart subtitle | read chart subtitle Aspose.Cells | Aspose.Cells Chart.SubTitle property | verify chart subtitle .NET | Aspose.Cells example C#
// Common Searches: how to change chart subtitle Aspose.Cells C# | read chart subtitle after setting Aspose.Cells | Aspose.Cells chart subtitle example | C# Aspose.Cells set subtitle on chart | verify chart subtitle value Aspose
// Developer Intent: Apply a custom subtitle to a chart and confirm the value programmatically.
// Use Cases: Add contextual information below a chart title for reports. | Log or display the subtitle text after setting it for UI validation. | Generate workbooks with customized chart subtitles for automated reporting.
// AI Prompts: Show C# code that sets a chart subtitle with Aspose.Cells and reads it back. | Provide an Aspose.Cells example that verifies a chart subtitle without saving the workbook. | Explain how to customize both the title and subtitle of a chart in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, inserts a column chart, assigns a main title and a custom subtitle, reads back the subtitle to verify the change, prints it, and saves the file as ChartWithSubtitle.xlsx using Aspose.Cells for C#.
class ChartSubtitleDemo
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

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Set the main title of the chart
        chart.Title.Text = "Main Chart Title";

        // Set a custom subtitle for the chart
        chart.SubTitle.Text = "Custom Chart Subtitle";

        // Verify that the subtitle was set correctly
        string subtitleText = chart.SubTitle.Text;
        Console.WriteLine("Subtitle set to: " + subtitleText);

        // Save the workbook to a file
        workbook.Save("ChartWithSubtitle.xlsx");
    }
}
