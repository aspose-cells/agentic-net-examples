// Title: Applying a dark theme to a timeline chart and saving it as PDF with Aspose.Cells for .NET
// Common Searches: c# aspnet apply dark background to chart area using Aspose.Cells | generate timeline line chart with dark theme and export to PDF in .NET | set axis colors and title font color for Aspose.Cells chart before PDF conversion | Aspose.Cells line chart dark mode example for PDF output | customize plot area background color in Aspose.Cells C#
// Tags: chart area color customization Aspose.Cells | timeline chart PDF generation Aspose.Cells | axis tick mark and color settings Aspose.Cells C# | plot area background styling Aspose.Cells | apply dark visual style to Aspose.Cells chart

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills it with date/value data, adds a line chart as a timeline, customizes chart and plot area colors for a dark theme, adjusts axis tick marks and title color for contrast, and saves the workbook as a PDF file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Populate sample data for the timeline (dates and values)
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value");
            DateTime startDate = new DateTime(2023, 1, 1);
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(startDate.AddDays(i));
                sheet.Cells[i + 1, 1].PutValue(i * 10);
            }

            // Add a line chart (used as a timeline substitute)
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
            var chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B11", true);
            chart.NSeries.CategoryData = "A2:A11";

            // Apply a dark theme to the chart area
            chart.ChartArea.Area.ForegroundColor = Color.FromArgb(30, 30, 30);
            chart.ChartArea.Area.BackgroundColor = Color.FromArgb(30, 30, 30);
            chart.ChartArea.Area.Transparency = 0;

            // Apply a dark background to the plot area
            chart.PlotArea.Area.ForegroundColor = Color.FromArgb(45, 45, 45);
            chart.PlotArea.Area.BackgroundColor = Color.FromArgb(45, 45, 45);

            // Configure axis appearance for better contrast
            var categoryAxis = chart.CategoryAxis;
            var valueAxis = chart.ValueAxis;

            // Tick marks
            categoryAxis.MajorTickMark = TickMarkType.Outside;
            categoryAxis.MinorTickMark = TickMarkType.Outside;
            valueAxis.MajorTickMark = TickMarkType.Outside;
            valueAxis.MinorTickMark = TickMarkType.Outside;

            // Set chart title with light color
            chart.Title.Text = "Dark Theme Timeline";
            chart.Title.Font.Color = Color.White;

            // Render the workbook to PDF
            string outputPath = "TimelineDarkTheme.pdf";
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
