// Title: Aspose.Cells for .NET – Enable Cell‑Based Data Labels (ShowCellRange) on a Chart Series
// Description: C# example that creates a workbook, adds a column chart, sets the series range, activates data labels with ShowCellRange, links them to cells C2:C3, applies blue font styling, and saves the file as SetShowCellRangeDemo.xlsx.
// Keywords: Aspose.Cells ShowCellRange | cell based data labels | chart series data labels .NET | LinkedSource Aspose.Cells | C# chart label from worksheet cells | Aspose.Cells chart formatting | Enable ShowCellRange property
// Common Searches: Aspose.Cells enable ShowCellRange for chart series | link chart data labels to worksheet cells Aspose.Cells C# | how to use LinkedSource with data labels in Aspose.Cells | customize chart data label font color Aspose.Cells | set ShowCellRange true Aspose.Cells example
// Developer Intent: Activate cell‑based data labels for a chart series and bind them to a worksheet range.
// Use Cases: Display custom text such as "100 units" from cells C2:C3 as data labels on a column chart. | Show both the numeric value and a linked cell’s text simultaneously on each data point. | Apply specific font styling (e.g., blue color) to data labels sourced from worksheet cells.
// AI Prompts: Generate C# code using Aspose.Cells that creates a line chart where each series has ShowCellRange enabled and links its data labels to a distinct cell range, applying red font color. | Provide an example that toggles ShowCellRange for multiple series in a single chart and assigns different LinkedSource ranges to each series. | Explain how to programmatically verify that ShowCellRange is enabled and retrieve the LinkedSource range for a series in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, adds a column chart, sets the series range, activates data labels with ShowCellRange, links them to cells C2:C3, applies blue font styling, and saves the file as SetShowCellRangeDemo.xlsx.
    public class SetShowCellRangeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["B3"].PutValue(200);
                sheet.Cells["C2"].PutValue("100 units");
                sheet.Cells["C3"].PutValue("200 units");

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series and categories
                chart.NSeries.Add("B2:B3", true);
                chart.NSeries.CategoryData = "A2:A3";

                // Access the first series
                Series series = chart.NSeries[0];

                // Enable data labels and activate cell‑based data labels
                series.DataLabels.ShowValue = true;          // optional, shows the numeric values
                series.DataLabels.ShowCellRange = true;      // activates cell range as data labels
                series.DataLabels.LinkedSource = "C2:C3";    // link to cells containing custom label text
                series.DataLabels.Font.Color = Color.Blue;  // optional styling

                // Save the workbook
                workbook.Save("SetShowCellRangeDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetShowCellRangeDemo.Run();
        }
    }
}
