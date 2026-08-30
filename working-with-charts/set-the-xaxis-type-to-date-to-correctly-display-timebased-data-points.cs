// Title: Configure a line chart X‑axis as a date/time scale with monthly major ticks and weekly minor ticks using Aspose.Cells for .NET (C#)
// AI Prompts: Create a line chart from a DateTime column and set the CategoryAxis to TimeScale with monthly major units in Aspose.Cells C#. | Apply a custom date format like "MMM dd, yyyy" to the X‑axis labels and define weekly minor tick intervals for a chart using Aspose.Cells .NET. | Show how to configure BaseUnitScale, MajorUnitScale, and MinorUnitScale for a date‑based chart axis in C# with Aspose.Cells.
// Common Searches: Aspose.Cells C# set chart X axis to date scale with monthly major ticks | How to display weekly minor tick marks on a line chart using Aspose.Cells .NET | Formatting X axis labels as month day year in Aspose.Cells line chart | TimeScale category axis configuration example for Aspose.Cells C#
// Tags: Aspose.Cells chart CategoryAxis TimeScale configuration | C# line chart date axis monthly major unit | Aspose.Cells set X axis minor unit days | custom X axis label format Aspose.Cells | Aspose.Cells chart date/time axis example

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a workbook, fills column A with DateTime values and column B with numeric data, adds a line chart, links the data ranges, and configures the chart's CategoryAxis to use a TimeScale. It sets monthly major ticks, weekly minor ticks, applies a custom date label format, and saves the file as SetXAxisToDateDemo.xlsx.
    public class SetXAxisToDateDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate worksheet with date‑based data
                sheet.Cells["A1"].PutValue("Date");
                sheet.Cells["B1"].PutValue("Value");

                sheet.Cells["A2"].PutValue(new DateTime(2024, 1, 1));
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue(new DateTime(2024, 2, 1));
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue(new DateTime(2024, 3, 1));
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["A5"].PutValue(new DateTime(2024, 4, 1));
                sheet.Cells["B5"].PutValue(25);

                // Add a line chart to display the data
                int chartIndex = sheet.Charts.Add(ChartType.Line, 6, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series (values) and the category (dates)
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Configure the X‑axis (category axis) to treat the data as dates
                chart.CategoryAxis.CategoryType = CategoryType.TimeScale;   // Date/Time axis
                chart.CategoryAxis.BaseUnitScale = TimeUnit.Months;        // Base unit (months)
                chart.CategoryAxis.MajorUnitScale = TimeUnit.Months;       // Major unit scale
                chart.CategoryAxis.MajorUnit = 1;                          // One month per major tick
                chart.CategoryAxis.MinorUnitScale = TimeUnit.Days;         // Minor unit scale
                chart.CategoryAxis.MinorUnit = 7;                          // One week per minor tick

                // Optional: format the X‑axis labels
                chart.NSeries[0].XValuesFormatCode = "mmm dd, yyyy";

                // Save the workbook
                workbook.Save("SetXAxisToDateDemo.xlsx");
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
            SetXAxisToDateDemo.Run();
        }
    }
}
