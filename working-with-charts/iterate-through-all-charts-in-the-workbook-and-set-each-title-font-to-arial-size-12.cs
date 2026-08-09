// Title: Set All Chart Title Fonts to Arial 12pt in Aspose.Cells (C#)
// Description: C# example that creates or loads a workbook, adds sample data and a chart, then iterates through every worksheet and chart to set each chart title font to Arial, size 12, and saves the file.
// Keywords: Aspose.Cells C# chart title font | set chart title Arial 12pt | iterate all charts workbook | bulk chart formatting Aspose | Excel chart title styling .NET | Aspose.Cells API title font | format chart titles programmatically | Excel workbook chart customization
// Common Searches: how to change font of all chart titles in Aspose.Cells C# | Aspose.Cells loop through worksheets to modify chart titles | set chart title font Arial for every chart in Excel using .NET | bulk update chart title style Aspose.Cells | C# code to format chart titles in a workbook
// Developer Intent: Apply a uniform Arial 12pt font to every chart title in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Enforce corporate branding by standardizing chart title fonts across generated reports. | Prepare Excel workbooks for publishing where all chart titles must follow a specific style. | Automate visual consistency when creating multiple charts programmatically.
// AI Prompts: Generate C# code with Aspose.Cells that changes all chart titles to Calibri 11pt. | Show how to hide chart titles that are empty while iterating through charts in a workbook. | Provide an example that sets the chart title color to blue for every chart using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// C# example that creates or loads a workbook, adds sample data and a chart, then iterates through every worksheet and chart to set each chart title font to Arial, size 12, and saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook (you can also load an existing one using new Workbook("file.xlsx"))
        Workbook workbook = new Workbook();

        // -------------------------------------------------
        // Sample data and a chart to demonstrate the logic
        // -------------------------------------------------
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);

        // Add a chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";
        chart.Title.Text = "Sample Chart";
        chart.Title.IsVisible = true;

        // -------------------------------------------------
        // Iterate through all charts in the workbook and set title font
        // -------------------------------------------------
        foreach (Worksheet ws in workbook.Worksheets)
        {
            foreach (Chart ch in ws.Charts)
            {
                // Set the title font to Arial, size 12
                ch.Title.Font.Name = "Arial";
                ch.Title.Font.Size = 12;
            }
        }

        // Save the workbook
        workbook.Save("AllChartTitlesArial.xlsx");
    }
}
