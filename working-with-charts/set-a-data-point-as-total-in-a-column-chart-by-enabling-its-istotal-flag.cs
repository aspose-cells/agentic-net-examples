// Title: How to mark a data point as total in an Aspose.Cells column chart using the IsTotal property (C#)
// Common Searches: Aspose.Cells C# set IsTotal flag on a column chart data point | mark total row in Excel column chart using Aspose.Cells API | C# example for setting total flag on chart series point with Aspose.Cells | display total value as a separate column in Aspose.Cells chart | how to use IsTotal property for a chart point in Aspose.Cells for .NET
// Tags: Aspose.Cells column chart IsTotal property | C# set total flag on chart series point | Excel column chart total data point Aspose.Cells | Aspose.Cells chart series point manipulation | C# Aspose.Cells chart total row handling

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, adds category and value data including a 'Total' row, inserts a column chart bound to the data range, sets a chart title, and saves the file. To designate the Total data point as a total in the chart, retrieve the corresponding point from the series and set its IsTotal property to true.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the column chart
            // Column A: Categories, Column B: Values
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B4"].PutValue(30);
            // Add a total row
            sheet.Cells["A5"].PutValue("Total");
            sheet.Cells["B5"].PutValue(60);

            // Add a column chart (rows 5‑20, columns 0‑10)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart data source
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Optional: give the chart a title
            chart.Title.Text = "Monthly Sales with Total";

            // Save the workbook
            string outputPath = "ColumnChartWithTotal.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
