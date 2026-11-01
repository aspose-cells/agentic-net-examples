// Title: Insert a 3‑D column chart into an Aspose.Cells worksheet with C# and set its data source and appearance
// AI Prompts: Create a 3‑D column chart on the first worksheet, bind the series to cells B2:C4, use A2:A4 for category labels, set the chart title to "Sales Overview", and apply a light gray plot‑area background using Aspose.Cells in C#. | Replace the existing chart data range with D2:E5, change the plot‑area background to LightBlue, and save the workbook as UpdatedChart.xlsx with Aspose.Cells C#.
// Common Searches: aspnet how to add a three dimensional column chart to a worksheet with Aspose.Cells | c# Aspose.Cells set chart series range and category axis for 3D column chart | changing plot area background color of a chart in Aspose.Cells .NET | save workbook with embedded 3D column chart using Aspose.Cells C# example
// Tags: Aspose.Cells create 3D column visualization | C# define chart data source Aspose.Cells | Aspose.Cells configure chart headings | Aspose.Cells modify chart visual style | Aspose.Cells save workbook with embedded chart

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.Drawing;

// The program creates a new workbook, populates sample data in A1:C4, inserts a 3‑D column chart covering rows 5‑25 and columns 0‑10, sets the title to "3D Column Chart", binds the series to B2:C4 with categories from A2:A4, applies a light‑gray plot‑area background, and saves the file as ThreeDColumnChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");

            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Insert a three‑dimensional column chart
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
            int chartIndex = sheet.Charts.Add(ChartType.Column3D, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "3D Column Chart";

            // Define the data range for the series (B2:C4) and enable categories
            chart.NSeries.Add("B2:C4", true);

            // Set the category (X‑axis) labels range
            chart.NSeries.CategoryData = "A2:A4";

            // Optional: format the chart appearance
            chart.PlotArea.Area.ForegroundColor = Color.LightGray;

            // Save the workbook with the chart embedded
            workbook.Save("ThreeDColumnChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
