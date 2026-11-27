// Title: How to set a currency number format with two decimal places for data labels on a bar chart using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set data label number format to currency on a bar chart | format bar chart data labels as $#,##0.00 in Aspose.Cells .NET | C# example applying custom currency format to chart data labels with Aspose.Cells | change data label font color and size in an Aspose.Cells bar chart | save a workbook containing a formatted bar chart as XLSX with Aspose.Cells
// Tags: chart data label currency formatting Aspose.Cells | custom number format string for series data labels C# | bar chart data label font styling Aspose.Cells | export workbook with chart to XLSX Aspose.Cells | apply number format to chart data labels .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SetDataLabelNumberFormat
{
    // The example creates a new workbook, adds a bar chart with fruit sales data, enables data labels, applies the custom currency format "$#,##0.00" (two decimal places) to the labels, customizes label font color and size, and saves the workbook as an XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Populate sample data for the bar chart
                // Column A: Categories
                // Column B: Values
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B2"].PutValue(1200.5);
                sheet.Cells["B3"].PutValue(850.75);
                sheet.Cells["B4"].PutValue(430.0);

                // Add a bar chart to the worksheet
                // Parameters: chart type, upper left row, upper left column, lower right row, lower right column
                int chartIndex = sheet.Charts.Add(ChartType.Bar, 6, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];
                chart.Title.Text = "Sales by Fruit";

                // Set the data source for the chart
                // NSeries.Add(dataRange, isVertical)
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Access the first series (there is only one)
                Series series = chart.NSeries[0];

                // Ensure data labels are displayed
                series.DataLabels.ShowValue = true;

                // Set the number format of the data labels to currency with two decimal places
                // Using a custom format string; the "$" symbol will adapt to the system locale if needed
                series.DataLabels.NumberFormat = "$#,##0.00";

                // Optionally, you can also set the font style for the data labels
                series.DataLabels.Font.Color = System.Drawing.Color.DarkBlue;
                series.DataLabels.Font.Size = 10;

                // Save the workbook to a file
                string outputPath = "BarChart_WithCurrencyDataLabels.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("Data label number format set to currency with two decimal places.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the chart:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
