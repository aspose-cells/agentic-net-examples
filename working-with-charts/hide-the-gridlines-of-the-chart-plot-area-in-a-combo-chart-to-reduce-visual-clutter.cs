// Title: Hide Gridlines in a Combo Chart with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample data, builds a column‑line combo chart, and disables major and minor gridlines on both value and category axes before saving as XLSX.
// Keywords: Aspose.Cells C# hide chart gridlines | combo chart gridlines Aspose.Cells | disable major gridlines Aspose.Cells | disable minor gridlines Aspose.Cells | Aspose.Cells chart axis visibility | C# Excel combo chart without gridlines | Aspose.Cells plot area styling
// Common Searches: Aspose.Cells hide gridlines combo chart | C# remove chart gridlines Aspose.Cells | disable major and minor gridlines in Excel chart using Aspose | hide category axis gridlines Aspose.Cells .NET | remove plot area gridlines Aspose.Cells chart
// Developer Intent: Hide all major and minor gridlines on the value and category axes of a combo chart to reduce visual clutter.
// Use Cases: Quarterly sales report with a column‑line combo chart that omits gridlines for a clean presentation. | Excel dashboard generation where multiple combo charts are rendered without gridlines to emphasize trends. | Financial statement export that follows corporate style guidelines by removing chart gridlines.
// AI Prompts: Generate C# code using Aspose.Cells that creates a combo chart and hides both major and minor gridlines on value and category axes. | Show an example of a column‑line combo chart in Aspose.Cells for .NET with all plot‑area gridlines disabled. | Explain how to toggle visibility of major and minor gridlines for specific axes in an Aspose.Cells chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsComboChartHideGridlines
{
    // Creates a workbook, adds sample data, builds a column‑line combo chart, and disables major and minor gridlines on both value and category axes before saving as XLSX.
    public class HideChartGridlines
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for a combo chart (column + line)
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Column Series");
                worksheet.Cells["C1"].PutValue("Line Series");

                worksheet.Cells["A2"].PutValue("Q1");
                worksheet.Cells["A3"].PutValue("Q2");
                worksheet.Cells["A4"].PutValue("Q3");
                worksheet.Cells["A5"].PutValue("Q4");

                worksheet.Cells["B2"].PutValue(120);
                worksheet.Cells["B3"].PutValue(150);
                worksheet.Cells["B4"].PutValue(180);
                worksheet.Cells["B5"].PutValue(210);

                worksheet.Cells["C2"].PutValue(80);
                worksheet.Cells["C3"].PutValue(130);
                worksheet.Cells["C4"].PutValue(170);
                worksheet.Cells["C5"].PutValue(200);

                // Add a chart (initially a column chart) which will become a combo chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the first series (column)
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries[0].Name = "Column Series";

                // Add a second series (line) and set its chart type to Line
                chart.NSeries.Add("C2:C5", true);
                chart.NSeries[1].Name = "Line Series";
                chart.NSeries[1].Type = ChartType.Line; // creates the combo effect

                // Set category (X) axis data
                chart.NSeries.CategoryData = "A2:A5";

                // Hide all gridlines in the plot area to reduce visual clutter
                chart.ValueAxis.MajorGridLines.IsVisible = false;
                chart.ValueAxis.MinorGridLines.IsVisible = false;
                chart.CategoryAxis.MajorGridLines.IsVisible = false;
                chart.CategoryAxis.MinorGridLines.IsVisible = false;

                // Save the workbook with the chart
                workbook.Save("ComboChart_HideGridlines.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved successfully as ComboChart_HideGridlines.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point required by the project
    public class Program
    {
        public static void Main(string[] args)
        {
            HideChartGridlines.Run();
        }
    }
}
