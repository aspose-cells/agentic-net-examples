// Title: Link chart data label number format to worksheet cells for dynamic formatting using Aspose.Cells in C#
// AI Prompts: Write C# code with Aspose.Cells that creates a column chart and sets DataLabels.LinkedSource to a cell range so the labels inherit the cells' number format. | Show how to enable NumberFormatLinked on a chart series in Aspose.Cells to make data labels automatically reflect formatting changes in the source cells. | Generate a complete Aspose.Cells example that populates raw values, formatted strings, adds a chart, and configures data labels to display the formatted strings from another column.
// Common Searches: Aspose.Cells C# chart data labels linked source range | set chart data label number format from worksheet cells Aspose.Cells | dynamic formatting of chart labels using LinkedSource property | inherit number format for chart series data labels Aspose.Cells example | bind chart data label text to cell values in Aspose.Cells
// Tags: Aspose.Cells chart data label linked source | NumberFormatLinked property Aspose.Cells | dynamic chart label formatting C# | column chart data labels from worksheet cells | inherit cell number format in chart labels

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsDynamicDataLabelFormatting
{
    // The program creates a workbook, fills category, raw numeric, and formatted string columns, adds a column chart, and configures the series so its data labels pull text from cells C2:C4 and inherit the number format of those cells via LinkedSource and NumberFormatLinked, then saves the file as DynamicDataLabelFormatting.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate category names
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            // Populate raw numeric values
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(1234.56);
            sheet.Cells["B3"].PutValue(7890.12);
            sheet.Cells["B4"].PutValue(3456.78);

            // Populate formatted strings that will drive the data label format
            // These cells contain the text representation with desired units
            sheet.Cells["C1"].PutValue("Formatted Value");
            sheet.Cells["C2"].PutValue("1,234.56 USD");
            sheet.Cells["C3"].PutValue("7,890.12 USD");
            sheet.Cells["C4"].PutValue("3,456.78 USD");

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (numeric values)
            chart.NSeries.Add("B2:B4", true);
            // Set the category (X‑axis) labels
            chart.NSeries.CategoryData = "A2:A4";

            // Access the first series
            Series series = chart.NSeries[0];

            // Enable data labels and link them to the formatted cells
            series.DataLabels.ShowValue = true;               // Show the value
            series.DataLabels.LinkedSource = "C2:C4";          // Source cells for label text
            series.DataLabels.NumberFormatLinked = true;      // Inherit number format from source cells

            // Optional: customize label appearance
            series.DataLabels.Font.Color = Color.DarkBlue;
            series.DataLabels.Font.Size = 10;

            // Save the workbook
            workbook.Save("DynamicDataLabelFormatting.xlsx");
        }
    }
}
