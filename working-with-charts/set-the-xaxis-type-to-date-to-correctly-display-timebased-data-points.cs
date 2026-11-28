// Title: Set the X‑Axis type to Date for a line chart using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells how to change chart X axis to date type | example of date axis in line chart with Aspose.Cells .NET | display time‑series data on X axis using Aspose.Cells chart API | Aspose.Cells set category axis to date format for line chart | create line chart with dates on X axis in C# using Aspose.Cells
// Tags: Aspose.Cells set chart X axis date type | line chart date axis C# | category axis date format Aspose.Cells | populate worksheet with date series Aspose.Cells | save workbook as xlsx Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartDateAxisDemo
{
    // This C# example creates a workbook, fills column A with sequential dates and column B with random numbers, adds a line chart, assigns the date range as the category (X) axis and the numeric range as the series values, sets the X‑axis to display dates, adds a chart title, and saves the file as an XLSX workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample time‑based data
                // Column A: Dates (starting from today, one day interval)
                // Column B: Random numeric values
                DateTime startDate = DateTime.Today;
                Random rnd = new Random();

                for (int i = 0; i < 10; i++)
                {
                    // Date value
                    Cell dateCell = cells[i, 0];
                    dateCell.PutValue(startDate.AddDays(i));

                    // Apply a date format so Excel recognises it as a date
                    Style dateStyle = dateCell.GetStyle();
                    dateStyle.Number = 14; // Built‑in date format (mm-dd-yy)
                    dateCell.SetStyle(dateStyle);

                    // Corresponding numeric value
                    cells[i, 1].PutValue(rnd.NextDouble() * 100);
                }

                // Add a line chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Line, 12, 0, 26, 7);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                // Category (X) axis: dates in column A
                // Values (Y) axis: numbers in column B
                chart.NSeries.Add("B1:B10", true);
                chart.NSeries.CategoryData = "A1:A10";

                // Optional: set a title for clarity
                chart.Title.Text = "Sample Date‑Based Line Chart";

                // Save the workbook to a file
                string outputPath = "DateAxisChartDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Report any errors that occur during processing
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
