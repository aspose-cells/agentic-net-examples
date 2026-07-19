// Title: Link Chart Series Data Labels to Source Cells in Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, populates raw and formatted values, adds a column chart with two series, and links each series' data‑label number format to its own formatted source range (D2:D4 and E2:E4). The NumberFormatLinked property keeps label formatting synchronized with the worksheet while custom font colors differentiate the series.
// Keywords: Aspose.Cells C# chart data labels | link data label format to cells | NumberFormatLinked property | column chart series formatting | custom font color Aspose.Cells | chart series linked source range | Aspose.Cells .NET example | dynamic chart label formatting
// Common Searches: Aspose.Cells link data label to source cells | C# chart data label number format linked source | How to use NumberFormatLinked in Aspose.Cells | Set custom font color for chart series labels Aspose.Cells | Create column chart with linked data labels .NET
// Developer Intent: Link each series' data‑label number format to its corresponding formatted source column.
// Use Cases: Display values with units on a chart by linking labels to pre‑formatted cells, ensuring consistency between worksheet and chart. | Automatically update chart label formatting when the source cells change, eliminating manual re‑formatting. | Apply distinct font colors to series labels while preserving linked number formats for clearer visual comparison.
// AI Prompts: Write C# code using Aspose.Cells to create a column chart where each series' data labels are linked to separate source ranges for number formatting and have custom font colors. | Explain the purpose of the NumberFormatLinked property in Aspose.Cells and show how to refresh linked labels after adding new rows to the data table. | Provide a step‑by‑step tutorial for linking data label formats to formatted cells for multiple chart series and customizing label appearance.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsExamples
{
    // This example creates a workbook, populates raw and formatted values, adds a column chart with two series, and links each series' data‑label number format to its own formatted source range (D2:D4 and E2:E4). The NumberFormatLinked property keeps label formatting synchronized with the worksheet while custom font colors differentiate the series.
    public class LinkSeriesDataLabelNumberFormat
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                // Column A: Category
                // Column B: Series 1 values
                // Column C: Series 2 values
                // Column D: Formatted values for Series 1 (e.g., with units)
                // Column E: Formatted values for Series 2
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["B3"].PutValue(200);
                sheet.Cells["B4"].PutValue(300);

                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(150);
                sheet.Cells["C3"].PutValue(250);
                sheet.Cells["C4"].PutValue(350);

                sheet.Cells["D1"].PutValue("Series1Formatted");
                sheet.Cells["D2"].PutValue("100 units");
                sheet.Cells["D3"].PutValue("200 units");
                sheet.Cells["D4"].PutValue("300 units");

                sheet.Cells["E1"].PutValue("Series2Formatted");
                sheet.Cells["E2"].PutValue("150 units");
                sheet.Cells["E3"].PutValue("250 units");
                sheet.Cells["E4"].PutValue("350 units");

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Add two series to the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries[0].Name = "Series 1";
                chart.NSeries.Add("C2:C4", true);
                chart.NSeries[1].Name = "Series 2";

                // Set category (X) data
                chart.NSeries.CategoryData = "A2:A4";

                // Configure data labels for each series
                // Series 1: link number format to D2:D4
                Series series1 = chart.NSeries[0];
                series1.DataLabels.ShowValue = true;
                series1.DataLabels.LinkedSource = "D2:D4";
                series1.DataLabels.NumberFormatLinked = true; // link format to source cells

                // Series 2: link number format to E2:E4
                Series series2 = chart.NSeries[1];
                series2.DataLabels.ShowValue = true;
                series2.DataLabels.LinkedSource = "E2:E4";
                series2.DataLabels.NumberFormatLinked = true; // link format to source cells

                // Optional: customize label appearance (e.g., font color)
                series1.DataLabels.Font.Color = Color.Blue;
                series2.DataLabels.Font.Color = Color.Green;

                // Save the workbook
                workbook.Save("LinkedSeriesDataLabels.xlsx");
                Console.WriteLine("Workbook saved successfully as LinkedSeriesDataLabels.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
