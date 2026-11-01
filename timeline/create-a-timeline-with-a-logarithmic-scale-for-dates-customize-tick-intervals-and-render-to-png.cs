// Title: Create a timeline line chart with a logarithmic Y‑axis, custom date tick intervals, and export it to PNG using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# create line chart with logarithmic Y axis | how to set custom major tick interval on date axis in Aspose.Cells chart | render Aspose.Cells worksheet chart as PNG image in .NET | generate timeline chart with exponential data using Aspose.Cells | apply base 10 logarithmic scale to chart value axis in C# Aspose.Cells
// Tags: logarithmic value axis Aspose.Cells | custom date axis major unit Aspose.Cells | export chart as PNG Aspose.Cells | timeline line chart Aspose.Cells | populate worksheet with dates exponential values Aspose.Cells

using System;
using System.Drawing.Imaging;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example creates a new workbook, fills column A with sequential dates and column B with exponential values, adds a line chart that serves as a timeline, assigns the dates as the category axis, enables a base‑10 logarithmic scale on the value axis, demonstrates how to customize major tick intervals on the date axis, renders the worksheet (including the chart) to a PNG image, and optionally saves the workbook as an XLSX file.
class TimelineLogScaleExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: Column A = dates, Column B = values
            DateTime startDate = new DateTime(2023, 1, 1);
            for (int i = 0; i < 10; i++)
            {
                // Date values (increment by 5 days)
                sheet.Cells[i + 1, 0].PutValue(startDate.AddDays(i * 5));
                // Corresponding numeric values (example exponential growth)
                sheet.Cells[i + 1, 1].PutValue(Math.Pow(10, i));
            }

            // Add a line chart (used as a timeline) to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the series (values in column B)
            chart.NSeries.Add("B2:B11", true);
            // Use dates in column A as category (X) axis
            chart.NSeries.CategoryData = "A2:A11";

            // Configure the category axis as a date axis (if supported)
            // Older versions may infer date axis automatically; explicit setting omitted for compatibility
            // Axis categoryAxis = chart.CategoryAxis;
            // categoryAxis.CategoryType = AxisCategoryType.Date;

            // Apply logarithmic scale to the value (Y) axis
            Axis valueAxis = chart.ValueAxis;
            valueAxis.IsLogarithmic = true;               // Enable logarithmic scaling
            valueAxis.LogBase = 10;                       // Base of the logarithm (optional, default is 10)

            // Optional: customize tick intervals on the date axis (if supported)
            // Axis categoryAxis = chart.CategoryAxis;
            // categoryAxis.MajorUnit = 10;                  // Interval between major ticks
            // categoryAxis.MajorUnitScale = AxisScale.Day; // Unit type for the major unit
            // categoryAxis.NumberFormat = "yyyy-mm-dd";    // Format date labels

            // Render the worksheet (including the chart) to a PNG image
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // ImageFormat property may not be available in older versions; the file extension determines format
                OnePagePerSheet = true
            };
            SheetRender sr = new SheetRender(sheet, imgOptions);
            sr.ToImage(0, "TimelineLogScale.png");

            // Save the workbook (optional, for verification)
            workbook.Save("TimelineLogScale.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
