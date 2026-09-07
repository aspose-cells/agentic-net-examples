// Title: Change the second data series to a line type in a combo chart using Aspose.Cells for .NET (C#)
// AI Prompts: Create a combo chart with a column series and then set the second series to a line chart using Aspose.Cells in C#. | Update an existing Aspose.Cells chart by changing the ChartType of a specific series to Line after the series has been added. | Add two data series to a workbook chart and convert the second series to a line series, optionally assigning it to a secondary axis with Aspose.Cells.
// Common Searches: Aspose.Cells C# change series type to line in a combo chart after creation | How to convert second series to line chart in Aspose.Cells .NET example | C# Aspose.Cells mixed column and line chart (combo) tutorial | Set secondary axis for line series in Aspose.Cells combo chart C#
// Tags: Aspose.Cells set series chart type | C# combo chart column line series | Aspose.Cells modify series after creation | Excel mixed chart Aspose.Cells .NET | Aspose.Cells secondary axis line series

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates creating a workbook, populating sample data, adding a combo chart with a column series, inserting two data series, converting the second series to a line type (and optionally assigning it to a secondary axis), and saving the workbook as an .xlsx file.
class ComboChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // A column: categories, B column: first series, C column: second series
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");

            string[] categories = { "Jan", "Feb", "Mar", "Apr", "May" };
            double[] series1 = { 10, 20, 30, 25, 15 };
            double[] series2 = { 5, 15, 25, 20, 10 };

            for (int i = 0; i < categories.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(categories[i]);   // A column
                sheet.Cells[i + 1, 1].PutValue(series1[i]);    // B column
                sheet.Cells[i + 1, 2].PutValue(series2[i]);    // C column
            }

            // Add a combo chart (initially a Column chart)
            // Position: from row 7, column 0 to row 25, column 7
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Combo Chart Example";

            // Add first series (Column) using data from B2:B6, categories from A2:A6
            chart.NSeries.Add("=Sheet1!$B$2:$B$6", true);
            // Add second series (will be changed to Line) using data from C2:C6
            chart.NSeries.Add("=Sheet1!$C$2:$C$6", true);

            // Change the second series type from Column to Line
            chart.NSeries[1].Type = ChartType.Line;

            // Optional: set the second series to use secondary axis if desired
            // chart.NSeries[1].IsSecondaryAxis = true;

            // Save the workbook
            string outputPath = "ComboChart_Output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
