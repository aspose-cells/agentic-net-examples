// Title: Aspose.Cells .NET – Link Chart Category Axis to Cell Range B2:B8 (C#)
// Description: Creates a workbook, fills B2:B8 with category labels and C2:C8 with values, adds a column chart, and binds the X‑axis to the B2:B8 range using NSeries.CategoryData before saving as an XLSX file.
// Keywords: Aspose.Cells chart category axis | C# NSeries CategoryData | link chart labels to cells | Aspose.Cells column chart example | bind chart categories range B2:B8 | .NET spreadsheet charting
// Common Searches: Aspose.Cells set chart category labels C# | NSeries.CategoryData usage .NET | bind chart axis to worksheet range Aspose.Cells | create column chart with custom categories Aspose.Cells | how to link chart categories to cells in C#
// Developer Intent: Bind the chart’s category (X‑axis) labels to the values in cells B2:B8.
// Use Cases: Generate a sales chart where product names stored in column B appear on the X‑axis. | Build a dynamic report that updates axis labels automatically when B2:B8 cells are edited. | Create a monthly performance chart that reads month names from a worksheet range for easy maintenance.
// AI Prompts: Show C# code to change the CategoryData range of an existing Aspose.Cells chart to a different column. | Provide an example that binds multiple series to a chart and assigns separate category ranges for each in Aspose.Cells. | Explain how to refresh chart category labels after modifying the source cells using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCategoryDataExample
{
    // Creates a workbook, fills B2:B8 with category labels and C2:C8 with values, adds a column chart, and binds the X‑axis to the B2:B8 range using NSeries.CategoryData before saving as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Column B will hold category labels (B2:B8)
            // Column C will hold numeric values for the series (C2:C8)
            sheet.Cells["B1"].PutValue("Category");
            sheet.Cells["C1"].PutValue("Value");
            for (int i = 2; i <= 8; i++)
            {
                sheet.Cells[$"B{i}"].PutValue($"Cat {i - 1}");
                sheet.Cells[$"C{i}"].PutValue(i * 10);
            }

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 10, 0, 30, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the series data range (values) – vertical series from C2:C8
            chart.NSeries.Add("C2:C8", true);

            // Link the category axis labels to the range B2:B8
            chart.NSeries.CategoryData = "B2:B8";

            // Save the workbook to a file
            workbook.Save("CategoryDataLinkedChart.xlsx", SaveFormat.Xlsx);
        }
    }
}
