// Title: How to set category axis tick labels to stacked (vertical) orientation using ChartTextDirectionType in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set category axis tick label direction to stacked | ChartTextDirectionType.Stacked example for vertical tick labels in .NET | Change tick label orientation of column chart using Aspose.Cells API | C# code to make chart axis labels vertical stacked in Excel file | Aspose.Cells chart axis label formatting stacked orientation tutorial
// Tags: ChartTextDirectionType stacked orientation | category axis tick label direction Aspose.Cells | vertical stacked tick labels for column chart | Aspose.Cells chart axis label formatting C# | export workbook with stacked tick labels to XLSX

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// The example creates a workbook, adds sample data and a column chart, sets the category axis tick labels to a stacked (vertical) orientation using ChartTextDirectionType.Stacked, applies a title, and saves the result as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories

            // Set the tick labels direction to Stacked (vertical orientation)
            chart.CategoryAxis.TickLabels.DirectionType = ChartTextDirectionType.Stacked;

            // Optional: set a title to illustrate the change
            chart.Title.Text = "Stacked Tick Labels Example";
            chart.Title.Font.Color = Color.Blue;

            // Save the workbook to a file
            string outputPath = "ChartTickLabelsStacked.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
