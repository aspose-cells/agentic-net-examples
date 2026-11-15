// Title: Create a column chart in an Aspose.Cells worksheet and bind its series to a numeric range using C#
// Common Searches: Aspose.Cells C# how to add a column chart and bind its series to cells A1:A5 | programmatically create a column chart in a .xlsx workbook with Aspose.Cells | binding chart NSeries to a numeric range in Aspose.Cells for .NET | clear default series and set data source for Aspose.Cells chart in C# | sample code for creating a column chart and linking it to data in Aspose.Cells
// Tags: Aspose.Cells column chart NSeries binding | C# set chart data source Aspose.Cells | create chart from numeric range Aspose.Cells | clear default series Aspose.Cells chart | bind chart series to cell range .xlsx C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new Workbook, renames the first worksheet, fills cells A1:A5 with numeric values, adds a column chart positioned on the sheet, clears any default series, binds a new series to the range A1:A5, names the series "Values", and saves the workbook as ChartExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // Populate numeric data in cells A1:A5
            double[] values = { 10, 20, 30, 40, 50 };
            for (int i = 0; i < values.Length; i++)
            {
                sheet.Cells[i, 0].PutValue(values[i]); // Column A
            }

            // Add a column chart to the worksheet (positioned from row 5, column 0 to row 20, column 10)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Sample Column Chart";

            // Clear any default series and bind a new series to the numeric range A1:A5
            chart.NSeries.Clear();
            chart.NSeries.Add("A1:A5", true);
            chart.NSeries[0].Name = "Values";

            // Save the workbook
            string outputPath = "ChartExample.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
