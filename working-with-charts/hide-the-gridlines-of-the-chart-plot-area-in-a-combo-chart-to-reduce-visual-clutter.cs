// Title: Aspose.Cells for .NET – Hide Gridlines in a Combo (Column & Line) Chart
// Description: Creates a workbook, adds sample data, builds a combo chart with column and line series, and disables major and minor gridlines on both value and category axes before saving the file as an XLSX workbook.
// Keywords: Aspose.Cells hide chart gridlines | combo chart gridlines .NET | disable major gridlines Aspose.Cells | remove minor gridlines chart | C# Aspose.Cells chart formatting | Excel combo chart without gridlines | chart axis visibility Aspose.Cells
// Common Searches: Aspose.Cells hide gridlines combo chart | C# remove chart gridlines Aspose.Cells | turn off major gridlines in Excel chart using Aspose | hide category axis gridlines Aspose.Cells | disable minor gridlines in combo chart .NET
// Developer Intent: Programmatically turn off all major and minor gridlines on both value and category axes of a combo (column‑plus‑line) chart using Aspose.Cells for .NET.
// Use Cases: Generate clean visual charts for dashboards or presentations by removing axis gridlines. | Automate Excel report creation where gridlines obscure data trends. | Prepare charts for PDF or image export with minimal visual clutter. | Apply consistent chart styling across multiple workbooks in a batch process.
// AI Prompts: Write C# code with Aspose.Cells to create a combo (column + line) chart and hide all major and minor gridlines. | Show how to set ValueAxis.MajorGridLines.IsVisible = false, ValueAxis.MinorGridLines.IsVisible = false, CategoryAxis.MajorGridLines.IsVisible = false, and CategoryAxis.MinorGridLines.IsVisible = false in Aspose.Cells. | Provide an example that adds sample data, builds a combo chart, customizes the title, disables gridlines on both axes, and saves the workbook as XLSX.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsComboChartGridlinesDemo
{
    // Creates a workbook, adds sample data, builds a combo chart with column and line series, and disables major and minor gridlines on both value and category axes before saving the file as an XLSX workbook.
    public class HideChartPlotAreaGridlines
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for a combo chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Column Series");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            worksheet.Cells["C1"].PutValue("Line Series");
            worksheet.Cells["C2"].PutValue(15);
            worksheet.Cells["C3"].PutValue(25);
            worksheet.Cells["C4"].PutValue(35);

            // Add a Combo chart (Column + Line)
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];
            chart.Type = ChartType.Column; // Primary chart type

            // Add column series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries[0].Name = "Column Series";

            // Add line series and set its chart type to Line
            chart.NSeries.Add("C2:C4", true);
            chart.NSeries[1].Name = "Line Series";
            chart.NSeries[1].Type = ChartType.Line;
            // Note: Secondary axis is optional; omitted for compatibility

            // Hide major and minor gridlines on both axes
            chart.ValueAxis.MajorGridLines.IsVisible = false;
            chart.ValueAxis.MinorGridLines.IsVisible = false;
            chart.CategoryAxis.MajorGridLines.IsVisible = false;
            chart.CategoryAxis.MinorGridLines.IsVisible = false;

            // Optional: customize chart title
            chart.Title.Text = "Combo Chart without Gridlines";

            // Save the workbook
            workbook.Save("ComboChart_NoGridlines.xlsx", SaveFormat.Xlsx);
        }
    }
}
