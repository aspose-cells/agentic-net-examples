// Title: C# Example – Set and Verify a Chart Subtitle Using Aspose.Cells for .NET
// Description: Shows how to create a workbook, add sample data, insert a column chart, set a custom string via the chart's SubTitle.Text property, read the subtitle back for verification, print it to the console, and save the workbook as ChartWithCustomSubtitle.xlsx.
// Keywords: Aspose.Cells chart subtitle C# | set chart subtitle .NET | read chart subtitle Aspose.Cells | verify chart subtitle property | Aspose.Cells SubTitle.Text | column chart subtitle example | Excel chart subtitle automation
// Common Searches: Aspose.Cells set chart subtitle C# | how to read chart subtitle Aspose.Cells | verify chart subtitle value .NET | chart subtitle example Aspose.Cells | C# code for chart subtitle in Excel
// Developer Intent: Programmatically assign a custom subtitle to a chart and confirm the assignment by reading the SubTitle.Text property.
// Use Cases: Add a descriptive subtitle to a generated column chart for clearer reporting. | Validate chart metadata in automated tests by setting and then reading the subtitle. | Create Excel dashboards where subtitles are dynamically generated and later extracted for analytics.
// AI Prompts: Generate C# code with Aspose.Cells that sets a chart's subtitle and then reads it back. | Explain how to verify a chart subtitle after assigning it in an Aspose.Cells workbook. | Show how to handle charts that lack a subtitle element before attempting to set SubTitle.Text.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSubtitleDemo
{
    // Shows how to create a workbook, add sample data, insert a column chart, set a custom string via the chart's SubTitle.Text property, read the subtitle back for verification, print it to the console, and save the workbook as ChartWithCustomSubtitle.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set the main title (optional)
            chart.Title.Text = "Main Chart Title";

            // Set the subtitle to a custom string
            chart.SubTitle.Text = "Custom Chart Subtitle";

            // Verify the subtitle by reading the property
            string retrievedSubtitle = chart.SubTitle.Text;
            Console.WriteLine("Subtitle set to: " + retrievedSubtitle);

            // Save the workbook to a file
            workbook.Save("ChartWithCustomSubtitle.xlsx");
        }
    }
}
