// Title: How to add a clickable hyperlink to a chart data label in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Create a column chart in a new workbook and set a hyperlink on the first data label so clicking it opens a web page. | Attach a hyperlink to the source cell of a specific chart point and verify that the corresponding data label acts as a clickable link. | Change the font color of a chart data label that contains a hyperlink using Aspose.Cells in C#.
// Common Searches: aspocells make chart label clickable link c# example | add hyperlink to specific series point in Aspose.Cells chart .NET | c# set chart data label to open external URL with Aspose.Cells | how to link Excel chart label to website using Aspose.Cells | Aspose.Cells column chart label hyperlink tutorial
// Tags: Aspose.Cells add hyperlink to chart label | C# chart point hyperlink | Excel column chart clickable label | format chart data label font color Aspose.Cells | worksheet hyperlink influencing chart label

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// The example creates a new workbook, fills cells A1:A5 and B1:B5 with sample data, adds a column chart, enables data labels, adds a hyperlink to cell B1 (making the first data label clickable and opening https://www.example.com), optionally changes the label's font color, and saves the file as ChartWithHyperlink.xlsx.
class AddHyperlinkToDataLabel
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category 1");
            sheet.Cells["A2"].PutValue("Category 2");
            sheet.Cells["A3"].PutValue("Category 3");
            sheet.Cells["A4"].PutValue("Category 4");
            sheet.Cells["A5"].PutValue("Category 5");

            sheet.Cells["B1"].PutValue(10);
            sheet.Cells["B2"].PutValue(20);
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["B4"].PutValue(40);
            sheet.Cells["B5"].PutValue(50);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B1:B5", true);
            chart.NSeries.CategoryData = "A1:A5";

            // Enable data labels to show values
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Add a hyperlink to the first data point's source cell (B1)
            // This effectively makes the data label act as a hyperlink when clicked.
            sheet.Hyperlinks.Add(0, 1, 1, 1, "https://www.example.com");

            // Optional: customize label appearance
            // chart.NSeries[0].DataLabels.Font.Color = Color.Blue;

            // Determine output file path
            string outputPath = "ChartWithHyperlink.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
