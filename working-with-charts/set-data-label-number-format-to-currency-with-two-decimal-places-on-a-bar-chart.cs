// Title: How to format bar chart data labels as currency with two decimal places using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a bar chart in an Excel workbook and applies the number format $#,##0.00 to its data labels with Aspose.Cells. | Show an example of enabling data labels on a chart series and setting a custom currency format in Aspose.Cells for .NET. | Provide a .NET snippet that adds sample data, inserts a bar chart, and formats the chart's data label values as currency with two decimal places.
// Common Searches: Aspose.Cells C# set chart data label format to currency with two decimals | How to display dollar amounts on bar chart labels using Aspose.Cells for .NET | Apply custom number format $#,##0.00 to Excel chart series data labels in C# | Create bar chart with currency formatted data labels in Aspose.Cells example
// Tags: Aspose.Cells chart data label currency number format | C# set chart series number format Aspose.Cells | Excel bar chart data label formatting .NET | custom number format for chart labels Aspose | currency display on chart values using Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a new workbook, fills cells A1:B4 with category and amount data, adds a bar chart, binds the amount range B2:B4 to the series, enables data labels, sets their number format to the currency pattern $#,##0.00, and saves the workbook as BarChartWithCurrencyLabels.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Amount");
            ws.Cells["A2"].PutValue("Jan");
            ws.Cells["A3"].PutValue("Feb");
            ws.Cells["A4"].PutValue("Mar");
            ws.Cells["B2"].PutValue(1200.5);
            ws.Cells["B3"].PutValue(850.75);
            ws.Cells["B4"].PutValue(430.0);

            // Add a bar chart (rows: 5-20, columns: 0-7)
            int chartIndex = ws.Charts.Add(ChartType.Bar, 5, 0, 20, 7);
            Chart chart = ws.Charts[chartIndex];

            // Set chart data range (values)
            chart.NSeries.Add("B2:B4", true);
            // Category data can be inferred; explicit setting omitted for compatibility

            // Show data labels with currency format
            chart.NSeries[0].DataLabels.ShowValue = true;
            chart.NSeries[0].DataLabels.NumberFormat = "$#,##0.00";

            // Define output file path
            string outputPath = "BarChartWithCurrencyLabels.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
