// Title: Generate three column charts without legends in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Create three column charts on a worksheet, assign each a title, hide the legend, and save the file with Aspose.Cells in C#. | Iterate to add column charts, bind data ranges, set ShowLegend = false, and export as XLSX using Aspose.Cells.
// Common Searches: how to hide chart legends when creating multiple charts with Aspose.Cells C# | aspnet generate series of column charts in one sheet using Aspose.Cells | c# loop to add charts to Excel workbook with Aspose.Cells | remove legend from Aspose.Cells chart programmatically | create slideshow-ready charts in Excel with Aspose.Cells .NET
// Tags: Aspose.Cells hide chart legend | Aspose.Cells create multiple column charts | Aspose.Cells chart loop generation | Aspose.Cells set chart title C# | Aspose.Cells save workbook with charts XLSX

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a new workbook, fills sample data, loops to add three column charts each with a title, disables their legends, and saves the workbook as ChartsPresentation.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and name it
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate sample data for the charts
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Series 1");
            dataSheet.Cells["C1"].PutValue("Series 2");
            dataSheet.Cells["A2"].PutValue("A");
            dataSheet.Cells["A3"].PutValue("B");
            dataSheet.Cells["A4"].PutValue("C");
            dataSheet.Cells["B2"].PutValue(10);
            dataSheet.Cells["B3"].PutValue(20);
            dataSheet.Cells["B4"].PutValue(30);
            dataSheet.Cells["C2"].PutValue(15);
            dataSheet.Cells["C3"].PutValue(25);
            dataSheet.Cells["C4"].PutValue(35);

            // Generate a series of charts (e.g., three charts) for a slideshow
            for (int i = 0; i < 3; i++)
            {
                // Add a new column chart positioned vertically with spacing between charts
                int chartIndex = dataSheet.Charts.Add(ChartType.Column, 5 + i * 15, 0, 20 + i * 15, 7);
                Chart chart = dataSheet.Charts[chartIndex];

                // Set a descriptive title for each chart
                chart.Title.Text = $"Chart {i + 1}";

                // Add data series (values) and categories to the chart
                chart.NSeries.Add("B2:C4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Hide the legend to reduce visual noise in the slideshow
                chart.ShowLegend = false;
            }

            // Save the workbook containing the charts
            workbook.Save("ChartsPresentation.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
