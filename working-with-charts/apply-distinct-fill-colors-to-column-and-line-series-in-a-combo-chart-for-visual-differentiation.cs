// Title: Set Different Fill Colors for Column and Line Series in a Combo Chart – Aspose.Cells for .NET (C#)
// Description: Creates a workbook with sales and profit data, adds a combo chart (column + line), converts the second series to a line type, applies a blue fill to the column series and a red border to the line series, then saves the file as ComboChartDistinctColors.xlsx.
// Keywords: Aspose.Cells | C# combo chart | column series fill color | line series border color | chart series color formatting | Excel chart styling .NET | Aspose.Cells ChartType.Column | Aspose.Cells ChartType.Line | custom chart colors | Excel export styling
// Common Searches: Aspose.Cells set column series color | change line series color in combo chart Aspose.Cells | C# Aspose.Cells combo chart color | how to format chart series colors Aspose.Cells | customize combo chart colors .NET
// Developer Intent: Apply distinct colors to the column and line series of a combo chart using Aspose.Cells.
// Use Cases: Sales vs. profit report where sales columns use a brand‑specific blue and profit line is highlighted in red for quick visual comparison. | Financial dashboard that meets accessibility contrast guidelines by assigning separate colors to each chart series. | Automated Excel export that consistently styles column series with a predefined fill and line series with a contrasting border.
// AI Prompts: Generate C# code with Aspose.Cells to create a combo chart, set the first series as a column with a custom fill color, and the second series as a line with a specific border color. | Explain how to use the Area.ForegroundColor and Border.Color properties to style column and line series in an Aspose.Cells chart. | Provide step‑by‑step instructions for applying brand colors to mixed chart types (column and line) in an Aspose.Cells workbook.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsComboChartDemo
{
    // Creates a workbook with sales and profit data, adds a combo chart (column + line), converts the second series to a line type, applies a blue fill to the column series and a red border to the line series, then saves the file as ComboChartDistinctColors.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                // Column A: Categories
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["A5"].PutValue("Apr");

                // Column B: Column series values
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);
                sheet.Cells["B5"].PutValue(200);

                // Column C: Line series values
                sheet.Cells["C1"].PutValue("Profit");
                sheet.Cells["C2"].PutValue(30);
                sheet.Cells["C3"].PutValue(45);
                sheet.Cells["C4"].PutValue(60);
                sheet.Cells["C5"].PutValue(80);

                // Add a combo chart (initially a column chart)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Add the column series (first series)
                chart.NSeries.Add("B2:B5", true);
                // Add the line series (second series)
                chart.NSeries.Add("C2:C5", true);

                // Set category (X‑axis) data
                chart.NSeries.CategoryData = "A2:A5";

                // Convert the second series to a line chart
                chart.NSeries[1].Type = ChartType.Line;

                // Apply distinct fill color to the column series
                chart.NSeries[0].Area.ForegroundColor = Color.FromArgb(79, 129, 189); // blue shade

                // Apply distinct line color to the line series (use Border for line color)
                chart.NSeries[1].Border.Color = Color.FromArgb(192, 80, 77); // red shade
                chart.NSeries[1].Border.IsVisible = true;

                // Save the workbook
                workbook.Save("ComboChartDistinctColors.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
