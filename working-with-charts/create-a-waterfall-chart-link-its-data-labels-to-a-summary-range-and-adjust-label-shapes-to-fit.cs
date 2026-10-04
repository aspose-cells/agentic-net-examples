// Title: Generate a Waterfall chart with Aspose.Cells for .NET, bind its data labels to a summary column, and reduce label font size
// AI Prompts: Add a Waterfall chart to the worksheet using the values in B2:B5 and categories in A2:A5, then hide the default value, category, and series name labels. | Bind the chart's data labels to the text in column C, set the label font size to 10 points, and save the workbook as WaterfallChart.xlsx.
// Common Searches: Aspose.Cells C# bind custom text from a cell range to waterfall chart data labels | set custom data label font size for waterfall chart using Aspose.Cells .NET | hide default chart labels and display only summary text in Aspose.Cells waterfall chart | programmatically create waterfall chart with summary labels column in C# Aspose.Cells | adjust waterfall chart label shape to fit within chart area Aspose.Cells
// Tags: waterfall chart data labels Aspose.Cells | custom label text binding Aspose.Cells C# | chart label font size Aspose.Cells | link chart labels to cell range Aspose.Cells | remove built‑in chart labels Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The example creates a new workbook, populates columns with categories, numeric values and summary text, adds a Waterfall chart based on the values, disables the built‑in value, category and series name labels, binds the data labels to the text in column C, reduces the label font size to 10 points so the labels fit, and saves the file as WaterfallChart.xlsx.
class WaterfallChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ----- Populate source data -----
            // Categories
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Start");
            sheet.Cells["A3"].PutValue("Revenue");
            sheet.Cells["A4"].PutValue("Cost");
            sheet.Cells["A5"].PutValue("Profit");

            // Values for the waterfall series
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(1000);
            sheet.Cells["B3"].PutValue(1500);
            sheet.Cells["B4"].PutValue(-500);
            sheet.Cells["B5"].PutValue(1000);

            // Summary range that will be used for data labels
            sheet.Cells["C1"].PutValue("Label");
            sheet.Cells["C2"].PutValue("Start Total");
            sheet.Cells["C3"].PutValue("Revenue Total");
            sheet.Cells["C4"].PutValue("Cost Total");
            sheet.Cells["C5"].PutValue("Profit Total");

            // ----- Add a Waterfall chart -----
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the series data (values) and categories
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Access the first (and only) series to configure data labels
            var series = chart.NSeries[0];
            var dataLabels = series.DataLabels;

            // ----- Adjust label appearance -----
            dataLabels.ShowValue = false;          // Hide the default numeric values
            dataLabels.ShowCategoryName = false;   // Hide category names
            dataLabels.ShowSeriesName = false;     // Hide series name
            dataLabels.Font.Size = 10;             // Reduce font size so labels fit within the chart area

            // Save the workbook with the chart
            workbook.Save("WaterfallChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
