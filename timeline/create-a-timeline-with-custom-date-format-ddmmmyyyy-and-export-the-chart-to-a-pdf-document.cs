// Title: Create a timeline line chart with dd-MMM-yyyy date formatting and export it to PDF using Aspose.Cells for .NET (C#)
// Common Searches: how to apply a custom dd-MMM-yyyy date style to cells for a timeline chart in Aspose.Cells C# | export a workbook that contains a line chart to PDF with Aspose.Cells .NET | generate a timeline chart using dates as categories in Aspose.Cells for C# | set a chart title and custom date format before saving as PDF in Aspose.Cells | save Aspose.Cells chart as PDF file programmatically in C#
// Tags: dd-MMM-yyyy number format Aspose.Cells | timeline line chart creation C# | export workbook with chart to PDF Aspose.Cells | date cell styling Aspose.Cells | chart title configuration Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, writes dates and numeric values, applies a dd-MMM-yyyy number format to the date cells, adds a line chart that uses the dates as the category axis (forming a timeline), sets a chart title, and saves the workbook with the chart as a PDF file using Aspose.Cells for .NET.
class TimelineChartToPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data: Column A with dates, Column B with values.
            DateTime[] dates = new DateTime[]
            {
                new DateTime(2023, 1, 1),
                new DateTime(2023, 2, 15),
                new DateTime(2023, 3, 30),
                new DateTime(2023, 5, 10),
                new DateTime(2023, 6, 25)
            };

            double[] values = { 10, 20, 15, 30, 25 };

            // Apply custom date format style.
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "dd-MMM-yyyy";
            StyleFlag dateFlag = new StyleFlag { NumberFormat = true };

            for (int i = 0; i < dates.Length; i++)
            {
                // Write date to column A.
                Cell dateCell = sheet.Cells[i, 0];
                dateCell.PutValue(dates[i]);
                dateCell.SetStyle(dateStyle, dateFlag);

                // Write corresponding value to column B.
                sheet.Cells[i, 1].PutValue(values[i]);
            }

            // Add a Line chart (used as a timeline representation).
            int upperLeftRow = 7;
            int upperLeftColumn = 0;
            int lowerRightRow = 25;
            int lowerRightColumn = 10;

            // Add returns the index of the newly created chart.
            int chartIndex = sheet.Charts.Add(ChartType.Line, upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn);
            Chart timelineChart = sheet.Charts[chartIndex];

            // Set the data range for the series (values column). Dates are used automatically as categories.
            timelineChart.NSeries.Add("B1:B5", true);

            // Optional: set a title for the chart.
            timelineChart.Title.Text = "Project Timeline";

            // Save the workbook (including the chart) as a PDF document.
            workbook.Save("TimelineChart.pdf", SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
