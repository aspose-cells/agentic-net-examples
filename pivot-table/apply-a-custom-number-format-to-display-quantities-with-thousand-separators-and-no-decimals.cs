using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsNumberFormatDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data (quantities) in column B
            worksheet.Cells["A1"].PutValue("Item");
            worksheet.Cells["A2"].PutValue("Alpha");
            worksheet.Cells["A3"].PutValue("Beta");
            worksheet.Cells["A4"].PutValue("Gamma");

            worksheet.Cells["B1"].PutValue("Quantity");
            worksheet.Cells["B2"].PutValue(1234);
            worksheet.Cells["B3"].PutValue(56789);
            worksheet.Cells["B4"].PutValue(101112);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;

            // Apply custom number format: thousand separators, no decimals
            // "#,##0" displays numbers like 1,234 or 56,789
            series.DataLabels.NumberFormat = "#,##0";

            // Save the workbook to a file
            workbook.Save("QuantitiesWithThousandSeparator.xlsx");
        }
    }
}