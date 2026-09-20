// Title: How to add an Arial 12‑point bold title to a column chart with Aspose.Cells in C#
// AI Prompts: Write C# code that creates a column chart using Aspose.Cells and applies an Arial 12‑point bold font to the chart title. | Show how to modify the Font properties of a chart title in an Aspose.Cells workbook programmatically. | Provide a step‑by‑step example of adding a custom styled title to an Excel chart with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set chart title font to Arial bold | example of formatting column chart title in Aspose.Cells .NET | how to apply 12 point Arial style to Excel chart title using Aspose.Cells
// Tags: Aspose.Cells set chart title font | C# column chart title styling | Excel chart title Arial bold | Aspose.Cells chart title formatting .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, inserts sample data, adds a column chart, and then sets the chart title text to "Sample Chart Title" with an Arial 12‑point bold font before saving the workbook as ChartWithLabel.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Add a label (chart title) with Arial font, size 12, bold style
            chart.Title.Text = "Sample Chart Title";
            chart.Title.Font.Name = "Arial";
            chart.Title.Font.Size = 12;
            chart.Title.Font.IsBold = true;

            // Save the workbook to a file
            workbook.Save("ChartWithLabel.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
