// Title: Set X‑Axis to Date (Time Scale) in Aspose.Cells for .NET
// Description: Creates a workbook, fills column A with DateTime values and column B with numbers, adds a line chart, and configures the CategoryAxis to use a time‑scale (date) axis with month‑based major units before saving the file.
// Keywords: Aspose.Cells X axis date | CategoryType.TimeScale | C# line chart date axis | chart base unit months | Aspose.Cells chart axis formatting
// Common Searches: Aspose.Cells set X axis to date | C# chart time scale Aspose.Cells | how to use CategoryType.TimeScale | month tick interval Aspose.Cells chart | display dates on Excel chart with Aspose.Cells
// Developer Intent: Configure a chart’s X‑axis as a date (time‑scale) axis and define month‑level tick intervals using Aspose.Cells in C#.
// Use Cases: Plot monthly sales data with dates on the X‑axis. | Create a project timeline where milestones are positioned by date. | Generate financial reports that show quarterly dates with proper spacing.
// AI Prompts: Show C# code to set the X‑axis to a time scale and set major units to months in Aspose.Cells. | How can I change an existing chart’s CategoryType to TimeScale and adjust tick spacing? | Explain how to format date labels after applying CategoryType.TimeScale in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills column A with DateTime values and column B with numbers, adds a line chart, and configures the CategoryAxis to use a time‑scale (date) axis with month‑based major units before saving the file.
    public class SetXAxisToDateDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate worksheet with date (X‑axis) and numeric (Y‑axis) data
                sheet.Cells["A1"].PutValue("Date");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue(new DateTime(2024, 1, 1));
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue(new DateTime(2024, 2, 1));
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue(new DateTime(2024, 3, 1));
                sheet.Cells["B4"].PutValue(30);

                // Add a line chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series and the category (X) axis
                chart.NSeries.Add("B2:B4", true);          // Y values
                chart.NSeries.CategoryData = "A2:A4";      // X values (dates)

                // Configure the X‑axis to treat values as dates (time scale)
                chart.CategoryAxis.CategoryType = CategoryType.TimeScale;

                // Define the base unit scale (months) for better tick spacing
                chart.CategoryAxis.BaseUnitScale = TimeUnit.Months;
                chart.CategoryAxis.MajorUnitScale = TimeUnit.Months;
                chart.CategoryAxis.MajorUnit = 1; // one month per major tick

                // Save the workbook
                string outputPath = "SetXAxisToDateDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
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
