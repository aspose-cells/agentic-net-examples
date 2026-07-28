// Title: Switch Chart X‑Axis to a Date (Time) Axis with Aspose.Cells for .NET (C#)
// Description: The sample creates a new workbook, writes DateTime values to column A and numbers to column B, adds a line chart, links the series and category ranges, then converts the X‑axis from a categorical axis to a time‑scale axis by setting CategoryAxis.CategoryType to CategoryType.TimeScale. It also demonstrates configuring BaseUnitScale, MajorUnitScale, MinorUnitScale for month and day tick intervals before saving the file as ChartWithDateAxis.xlsx.
// Keywords: Aspose.Cells | C# | .NET | chart X axis date | time scale axis | CategoryType.TimeScale | line chart | date axis scaling | BaseUnitScale | MajorUnitScale | MinorUnitScale | Excel chart automation
// Common Searches: Aspose.Cells change X axis to date axis C# | time scale axis example Aspose.Cells .NET | set BaseUnitScale and MajorUnitScale for chart axis Aspose.Cells | convert categorical axis to time axis Aspose.Cells chart | date axis tick interval Aspose.Cells line chart
// Developer Intent: Convert a chart's X‑axis from categorical to date (time) scale.
// Use Cases: Plot monthly sales data where the X‑axis shows actual month dates. | Create a financial trend report with daily dates on the X‑axis and weekly minor ticks. | Generate a project timeline spreadsheet that accurately reflects irregular time intervals using a time‑scale axis.
// AI Prompts: Write C# code using Aspose.Cells to add a line chart and switch its X‑axis to a time‑scale axis with monthly major ticks. | Explain how to set BaseUnitScale, MajorUnitScale, and MinorUnitScale for a date axis in Aspose.Cells charts. | Provide step‑by‑step instructions to change a chart's category axis to a date axis and customize tick intervals in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The sample creates a new workbook, writes DateTime values to column A and numbers to column B, adds a line chart, links the series and category ranges, then converts the X‑axis from a categorical axis to a time‑scale axis by setting CategoryAxis.CategoryType to CategoryType.TimeScale. It also demonstrates configuring BaseUnitScale, MajorUnitScale, MinorUnitScale for month and day tick intervals before saving the file as ChartWithDateAxis.xlsx.
    public class SwitchXAxisToDateAxis
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data: dates in column A, values in column B
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

                // Add a line chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Line, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series (values) and the category (dates)
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Switch the X axis from a categorical axis to a date (time) axis
                chart.CategoryAxis.CategoryType = CategoryType.TimeScale;

                // Optional: define the base unit and major/minor units for better scaling
                chart.CategoryAxis.BaseUnitScale = TimeUnit.Months;   // Base unit (months)
                chart.CategoryAxis.MajorUnitScale = TimeUnit.Months; // Major tick every month
                chart.CategoryAxis.MajorUnit = 1;
                chart.CategoryAxis.MinorUnitScale = TimeUnit.Days;   // Minor tick every day
                chart.CategoryAxis.MinorUnit = 7;

                // Save the workbook with the configured chart
                string outputPath = "ChartWithDateAxis.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            SwitchXAxisToDateAxis.Run();
        }
    }
}
