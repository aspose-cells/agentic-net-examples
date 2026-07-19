// Title: Aspose.Cells for .NET – Set and Verify a Chart Subtitle (C#)
// Description: Creates a workbook, adds a column chart with sample data, assigns a custom string to the chart's SubTitle.Text property, reads the subtitle back to confirm the change, outputs it to the console, and saves the file as ChartWithSubtitle.xlsx.
// Keywords: Aspose.Cells chart subtitle | C# set chart subtitle | read chart subtitle Aspose.Cells | verify chart subtitle .NET | Aspose.Cells SubTitle.Text example
// Common Searches: how to change chart subtitle using Aspose.Cells | read chart subtitle after setting it in .NET | Aspose.Cells verify subtitle value | C# Aspose.Cells chart subtitle code
// Developer Intent: Programmatically assign a custom subtitle to a chart and confirm the assignment by reading the property.
// Use Cases: Add a descriptive subtitle to a generated Excel chart for clearer reporting. | Ensure the subtitle matches dynamic content before finalizing the workbook. | Automate subtitle updates based on worksheet data in batch Excel generation.
// AI Prompts: Write C# code with Aspose.Cells that sets a chart subtitle, reads it back, and prints the result. | Show how to conditionally modify a chart subtitle based on cell values using Aspose.Cells for .NET. | Provide an example that validates a chart subtitle after changing it with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSubtitleDemo
{
    // Creates a workbook, adds a column chart with sample data, assigns a custom string to the chart's SubTitle.Text property, reads the subtitle back to confirm the change, outputs it to the console, and saves the file as ChartWithSubtitle.xlsx.
    public class Program
    {
        public static void Main()
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

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart data range
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set main title (optional)
            chart.Title.Text = "Main Chart Title";

            // Set custom subtitle
            chart.SubTitle.Text = "Custom Chart Subtitle";

            // Verify the subtitle by reading the property
            string subtitle = chart.SubTitle.Text;
            Console.WriteLine("Subtitle set to: " + subtitle);

            // Save the workbook
            workbook.Save("ChartWithSubtitle.xlsx");
        }
    }
}
