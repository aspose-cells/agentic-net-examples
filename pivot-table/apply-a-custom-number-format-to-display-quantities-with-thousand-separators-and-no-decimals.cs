// Title: How to apply a custom number format with thousand separators to chart data labels in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a column chart with Aspose.Cells and sets the data label number format to "#,##0" to show values with commas and no decimal places. | Show the steps to enable data labels on an Aspose.Cells chart series and apply a thousand‑separator format without decimals. | Provide an example of formatting chart series values using a custom number format in an Aspose.Cells workbook.
// Common Searches: Aspose.Cells C# chart data label format "#,##0" thousand separator | set custom number format for Excel chart values using Aspose.Cells | how to display numbers with commas in Aspose.Cells chart labels | C# Aspose.Cells column chart label formatting without decimals | apply number format to series data labels in Aspose.Cells workbook
// Tags: chart data label number format Aspose.Cells | thousand separator formatting Excel chart C# | custom number format "#,##0" Aspose.Cells | apply number format to series values Aspose.Cells | column chart label formatting Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsNumberFormatDemo
{
    // Creates a workbook, adds item and quantity data, inserts a column chart, enables data labels, sets the data label NumberFormat to "#,##0" to display quantities with commas and no decimal places, and saves the file as QuantitiesWithThousandSeparator.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data (quantities)
            worksheet.Cells["A1"].PutValue("Item");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["A3"].PutValue("Banana");
            worksheet.Cells["A4"].PutValue("Cherry");

            worksheet.Cells["B1"].PutValue("Quantity");
            worksheet.Cells["B2"].PutValue(1234);
            worksheet.Cells["B3"].PutValue(56789);
            worksheet.Cells["B4"].PutValue(101112);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";     // Categories

            // Enable data labels to show the values on the chart
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;

            // Apply custom number format: thousand separators, no decimals
            series.DataLabels.NumberFormat = "#,##0";

            // Save the workbook
            workbook.Save("QuantitiesWithThousandSeparator.xlsx");
        }
    }
}
