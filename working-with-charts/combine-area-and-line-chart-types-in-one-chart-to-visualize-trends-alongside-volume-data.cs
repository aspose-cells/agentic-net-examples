// Title: Create a Combo Area and Line Chart with Aspose.Cells for .NET (C#)
// Description: This example builds a new workbook, fills columns with month, volume, and price data, adds an Area chart, inserts two series, sets the month range as the category axis, converts the second series to a Line type, optionally styles the line, and saves the result as ComboAreaLineChart.xlsx.
// Keywords: Aspose.Cells | C# chart example | combo chart | area chart | line chart | combined chart types | Excel chart series | ChartType.Area | ChartType.Line | visualize volume and price | chart formatting Aspose.Cells
// Common Searches: Aspose.Cells create combo area line chart C# | how to add line series to an area chart using Aspose.Cells | change series type to line after creating area chart Aspose.Cells | combined chart types in Excel with Aspose.Cells .NET | set different chart types for individual series Aspose.Cells
// Developer Intent: Generate a single Excel chart that shows volume as an area series and price as a line series.
// Use Cases: Financial reports that compare monthly sales volume (area) with average selling price (line). | Manufacturing dashboards displaying production quantity (area) alongside unit cost trend (line). | Marketing analytics sheets that overlay website traffic (area) with conversion rate (line).
// AI Prompts: Write C# code using Aspose.Cells to create a combo chart where column B is plotted as an Area series and column C as a Line series, using column A for categories. | Explain how to switch the chart type of a specific series to Line after adding an Area chart in Aspose.Cells. | Provide code to set the line series color to red and weight to 0.75 in a combined Area‑Line chart with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;   // Required for chart drawing objects

// This example builds a new workbook, fills columns with month, volume, and price data, adds an Area chart, inserts two series, sets the month range as the category axis, converts the second series to a Line type, optionally styles the line, and saves the result as ComboAreaLineChart.xlsx.
class ComboAreaLineChart
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: Month, Volume (for Area), Price (for Line)
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Volume");
            sheet.Cells["C1"].PutValue("Price");

            string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
            int[] volume = { 120, 150, 180, 130, 170 };
            double[] price = { 10.5, 12.0, 11.8, 13.2, 12.5 };

            for (int i = 0; i < months.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(months[i]);   // Column A: Month
                sheet.Cells[i + 1, 1].PutValue(volume[i]);   // Column B: Volume
                sheet.Cells[i + 1, 2].PutValue(price[i]);    // Column C: Price
            }

            // Add an Area chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Area, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Volume (Area) & Price (Line)";

            // Add two series: Volume (Area) and Price (Line)
            chart.NSeries.Add("B2:B6", true); // Volume series
            chart.NSeries.Add("C2:C6", true); // Price series (will be changed to Line)

            // Set the category (X‑axis) data – the months
            chart.NSeries.CategoryData = "A2:A6";

            // Convert the second series to a Line chart type
            chart.NSeries[1].Type = ChartType.Line;

            // Optional formatting for the line series (if supported by the library version)
            // Uncomment the following lines if the Series.Line property is available:
            // chart.NSeries[1].Line.Weight = 0.75;
            // chart.NSeries[1].Line.Color = Color.Red;

            // Save the workbook with the combined chart
            workbook.Save("ComboAreaLineChart.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
