// Title: Bold and Centered Chart Data Labels with Aspose.Cells for .NET
// Description: Shows how to create a workbook, add a column chart, enable data labels, and format them so the text appears in a bold typeface and is horizontally centered using the Aspose.Cells API.
// Keywords: Aspose.Cells C# chart data label formatting | bold data labels Aspose.Cells | center align chart labels .NET | Excel chart label style programmatically | apply font style to chart data labels | Aspose.Cells chart customization
// Common Searches: C# set chart data label to bold Aspose.Cells | center data labels in column chart Aspose.Cells | how to format chart labels programmatically with Aspose.Cells | Aspose.Cells change font of data labels | Excel chart label alignment using Aspose.Cells .NET
// Developer Intent: The developer wants chart data labels to be displayed in bold typeface and horizontally centered for improved readability.
// Use Cases: Create a quarterly sales column chart where each column shows its value in bold, centered labels for presentation decks. | Generate a budgeting workbook that highlights expense categories with bold, centered chart labels. | Automate performance dashboards that require uniformly styled data labels across multiple series.
// AI Prompts: Provide C# code to also set the font size and color of chart data labels while keeping them bold and centered. | Show how to apply italic style and right‑alignment to data labels on a line chart using Aspose.Cells. | Explain how to format data labels for all series in a multi‑series chart with the same bold, centered appearance.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, add a column chart, enable data labels, and format them so the text appears in a bold typeface and is horizontally centered using the Aspose.Cells API.
    public class DataLabelsBoldCenteredDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["A5"].PutValue("D");

                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);
                worksheet.Cells["B5"].PutValue(40);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Enable data labels and format them
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;
                series.DataLabels.Font.IsBold = true;
                series.DataLabels.TextHorizontalAlignment = TextAlignmentType.Center;
                series.DataLabels.ApplyFont();

                // Save the workbook
                workbook.Save("DataLabelsBoldCenteredDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            DataLabelsBoldCenteredDemo.Run();
        }
    }
}
