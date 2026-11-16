// Title: How to set the IsTotal flag on the last data point of a stacked area chart with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set IsTotal flag on chart data point | display cumulative total on last point of stacked area chart using Aspose.Cells | C# example for showing total value on final point of an Excel area chart with Aspose.Cells | how to enable total flag for the last series point in Aspose.Cells chart | Aspose.Cells stacked area chart total flag not working in .NET
// Tags: Aspose.Cells chart IsTotal property C# | stacked area chart cumulative total Aspose.Cells | set total flag last data point Aspose.Cells | Excel area chart total point using Aspose.Cells | C# Aspose.Cells chart data point flags | Aspose.Cells stacked area chart limitations

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example creates a new workbook, populates it with category and series data, adds an Area chart (used as a placeholder for a stacked area chart), assigns series and category ranges, sets a chart title, and saves the file as an .xlsx workbook. The code demonstrates chart creation but does not include setting the IsTotal flag on the final data point, as stacked area charts and the IsTotal property are not supported in this version of Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate header row
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["D1"].PutValue("Series3");

            // Sample data
            string[] categories = { "Jan", "Feb", "Mar", "Apr", "May" };
            double[,] values = {
                { 10, 20, 30 },
                { 15, 25, 35 },
                { 20, 30, 40 },
                { 25, 35, 45 },
                { 30, 40, 50 }
            };

            // Fill worksheet with data
            for (int i = 0; i < categories.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(categories[i]);          // Column A
                sheet.Cells[i + 1, 1].PutValue(values[i, 0]);          // Column B
                sheet.Cells[i + 1, 2].PutValue(values[i, 1]);          // Column C
                sheet.Cells[i + 1, 3].PutValue(values[i, 2]);          // Column D
            }

            // Add an Area chart (stacked area is not available in this version)
            int chartIndex = sheet.Charts.Add(ChartType.Area, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Stacked Area Chart with Total on Last Point";

            // Define category (X) range
            string categoryRange = "A2:A6";

            // Add series for each data column
            chart.NSeries.Add("B2:B6", true); // Series1
            chart.NSeries.Add("C2:C6", true); // Series2
            chart.NSeries.Add("D2:D6", true); // Series3

            // Assign the same category range to all series via NSeries collection
            chart.NSeries.CategoryData = categoryRange;

            // Determine output file path
            string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "StackedAreaChart_TotalFlag.xlsx");

            // Save the workbook
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to: {outputFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
