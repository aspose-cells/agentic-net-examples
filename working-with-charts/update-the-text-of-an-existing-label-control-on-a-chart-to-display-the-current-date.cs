// Title: Set chart data labels to today's date with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a column chart, populates sample data, enables data labels, disables auto‑generated text for each point, assigns the current date (yyyy‑MM‑dd) to the label, and saves the file.
// Keywords: Aspose.Cells chart data label | C# set chart label text | current date label Aspose.Cells | custom chart point labels .NET | DataLabels.Text Aspose.Cells | disable auto text chart label | update chart labels programmatically
// Common Searches: Aspose.Cells set custom text for chart data labels | C# change chart point label to today’s date | How to disable auto text on Aspose.Cells chart labels | Update column chart labels with dynamic date Aspose.Cells | Aspose.Cells DataLabels.Text example
// Developer Intent: Programmatically replace each chart point’s data label with the current date.
// Use Cases: Add a generation timestamp to every column in a sales chart for audit trails. | Display the report date on KPI chart labels so viewers know data freshness. | Insert a daily update date into financial chart labels to indicate when values were captured.
// AI Prompts: Generate C# code using Aspose.Cells that disables auto text for chart points and sets DataLabels.Text to DateTime.Now formatted as yyyy‑MM‑dd. | Show how to update only the first series of a line chart with today’s date as the label text in Aspose.Cells. | Explain how to change the DataLabels.Text property for each ChartPoint while keeping ShowValue enabled.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLabelUpdateDemo
{
    // Creates a workbook, adds a column chart, populates sample data, enables data labels, disables auto‑generated text for each point, assigns the current date (yyyy‑MM‑dd) to the label, and saves the file.
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

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the series
            foreach (Series series in chart.NSeries)
            {
                series.DataLabels.ShowValue = true;

                // Update each data label to display the current date
                foreach (ChartPoint point in series.Points)
                {
                    // Disable auto-generated text so we can set custom text
                    point.DataLabels.IsAutoText = false;

                    // Set the label text to today's date (e.g., "2026-08-10")
                    point.DataLabels.Text = DateTime.Now.ToString("yyyy-MM-dd");
                }
            }

            // Save the workbook with the updated chart labels
            workbook.Save("ChartWithDateLabels.xlsx");
        }
    }
}
