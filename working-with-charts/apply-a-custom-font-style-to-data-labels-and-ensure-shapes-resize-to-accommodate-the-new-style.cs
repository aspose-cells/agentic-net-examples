// Title: Apply a custom font to chart data labels and auto‑resize label shapes using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that sets the font name, size, color, bold and italic attributes for data labels on an Aspose.Cells column chart. | Show how to make chart data label shapes automatically adjust their size after changing the label font in Aspose.Cells. | Create a complete example that builds a workbook, adds a column chart, enables styled data labels, and saves the file as an .xlsx.
// Common Searches: Aspose.Cells C# change font of Excel chart data labels | how to auto adjust data label size after font change in Aspose.Cells chart | set bold italic font for column chart data labels using Aspose.Cells .NET | C# example for styling chart data labels and saving workbook with Aspose.Cells | resize chart data label shape to fit custom font Aspose.Cells
// Tags: set chart data label font Aspose.Cells | auto resize data label shape .NET | column chart label styling C# | Aspose.Cells custom font for Excel chart labels | chart data label formatting with Aspose.Cells | Excel chart label auto size adjustment C#

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, adds a column chart with sample data, enables data labels, applies an Arial Black 14‑point blue bold‑italic font to the labels, and saves the file as CustomFontDataLabels.xlsx, allowing the label shapes to automatically resize to fit the new style.
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
            sheet.Name = "Data";

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Sample Chart";

            // Add a series to the chart (values range)
            int seriesIndex = chart.NSeries.Add("B2:B4", true);

            // Optionally set category (X) axis data if needed
            // chart.NSeries[seriesIndex].CategoryData = "A2:A4";

            // Show data labels (values) on the series
            DataLabels dataLabel = chart.NSeries[seriesIndex].DataLabels;
            dataLabel.ShowValue = true; // display the value

            // Apply custom font style to the data labels
            dataLabel.Font.Name = "Arial Black";
            dataLabel.Font.Size = 14;
            dataLabel.Font.Color = Color.Blue;
            dataLabel.Font.IsBold = true;
            dataLabel.Font.IsItalic = true;

            // Save the workbook to a file
            string outputPath = "CustomFontDataLabels.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
