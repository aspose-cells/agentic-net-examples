// Title: Combine Area and Line Chart in Aspose.Cells for .NET – Price Trend with Volume
// Description: Creates a new workbook, fills it with dates, price, and volume data, then adds an Area series for price and a Line series for volume on the same chart. Includes optional secondary‑axis configuration, axis titles, chart title, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells C# mixed chart | area chart with line series | combine chart types .NET | change series type Aspose.Cells | secondary axis line series | price and volume Excel chart | financial chart Aspose.Cells | multiple series chart C# | Excel area‑line chart example
// Common Searches: Aspose.Cells add line series to area chart | mixed chart example Aspose.Cells C# | how to set secondary axis in Aspose.Cells chart | change chart type of a series Aspose.Cells | price volume chart Aspose.Cells tutorial
// Developer Intent: Generate an Excel file that contains a single chart with an area series for price and a line series for volume, optionally using a secondary axis.
// Use Cases: Display stock price as a filled area while showing daily trading volume as a line on the same chart. | Build financial dashboards that compare a trend metric with a count metric using different visual styles. | Create periodic reports that overlay a KPI trend with related volume data in a mixed chart.
// AI Prompts: Write C# code with Aspose.Cells to add an area series and a line series to one chart, placing the line series on a secondary axis. | Explain how to modify the chart type of an existing series after it has been added in Aspose.Cells. | Provide steps to format the date category axis, set chart and axis titles, and enable a secondary axis for a mixed area‑line chart in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a new workbook, fills it with dates, price, and volume data, then adds an Area series for price and a Line series for volume on the same chart. Includes optional secondary‑axis configuration, axis titles, chart title, and saves the file as an Excel workbook.
class CombineAreaAndLineChart
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: Date, Price (trend), Volume (bars)
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["C1"].PutValue("Volume");

            for (int i = 2; i <= 11; i++)
            {
                // Use OADate for Excel date values
                sheet.Cells[i - 1, 0].PutValue(DateTime.Today.AddDays(i - 2).ToOADate());
                sheet.Cells[i - 1, 1].PutValue(i * 10);      // Price
                sheet.Cells[i - 1, 2].PutValue(i * 1000);   // Volume
            }

            // Add an Area chart (price) and a Line series (volume)
            int chartIndex = sheet.Charts.Add(ChartType.Area, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Price series – Area
            chart.NSeries.Add("B2:B11", true);
            chart.NSeries[0].Name = "Price";

            // Volume series – Line
            chart.NSeries.Add("C2:C11", true);
            chart.NSeries[1].Name = "Volume";
            chart.NSeries[1].Type = ChartType.Line;   // Change series type to Line

            // Optional: If secondary axis is needed, uncomment the following lines
            // chart.SecondaryValueAxis.IsVisible = true;
            // chart.NSeries[1].IsOnSecondaryAxis = true;
            // chart.SecondaryValueAxis.Title.Text = "Volume";

            // Set chart titles and axis titles
            chart.Title.Text = "Price Trend (Area) and Volume (Line)";
            chart.CategoryAxis.Title.Text = "Date";
            chart.ValueAxis.Title.Text = "Price";

            // Determine output file path
            string outputPath = "CombinedAreaLineChart.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
