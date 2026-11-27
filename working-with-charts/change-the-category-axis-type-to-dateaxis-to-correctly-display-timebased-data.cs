// Title: Set a line chart’s category axis to a Date (TimeScale) axis with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart category axis to TimeScale | Create line chart with monthly dates on X axis using Aspose.Cells | Programmatically configure date based X axis in an Excel chart with Aspose.Cells | Adjust BaseUnitScale to months for date axis in Aspose.Cells .NET
// Tags: Aspose.Cells chart category axis Date type | C# line chart date axis implementation | TimeScale axis configuration Aspose.Cells | Excel chart X axis date handling Aspose.Cells | Configure month tick spacing Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills column A with monthly dates starting January 2023 and column B with sales figures, adds a line chart, links the data ranges, sets the category axis type to TimeScale to treat the X‑axis as dates, optionally sets the BaseUnitScale to months for appropriate tick intervals, and saves the file as DateAxisChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate the worksheet with date‑based data
            worksheet.Cells["A1"].PutValue("Date");
            worksheet.Cells["B1"].PutValue("Sales");

            DateTime startDate = new DateTime(2023, 1, 1);
            for (int i = 0; i < 12; i++)
            {
                // Column A – dates (one per month)
                worksheet.Cells[i + 1, 0].PutValue(startDate.AddMonths(i));
                // Column B – some numeric values
                worksheet.Cells[i + 1, 1].PutValue(1000 + i * 200);
            }

            // Add a line chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Line, 5, 0, 25, 10);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the series and the category (X) axis
            chart.NSeries.Add("B2:B13", true);          // Values
            chart.NSeries.CategoryData = "A2:A13";     // Dates

            // Use TimeScale for date‑based category axis
            chart.CategoryAxis.CategoryType = CategoryType.TimeScale;

            // Optional: define the base unit scale (e.g., months) for better tick spacing
            chart.CategoryAxis.BaseUnitScale = TimeUnit.Months;

            // Save the workbook
            string outputPath = "DateAxisChart.xlsx";
            try
            {
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine("Error saving workbook: " + saveEx.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
