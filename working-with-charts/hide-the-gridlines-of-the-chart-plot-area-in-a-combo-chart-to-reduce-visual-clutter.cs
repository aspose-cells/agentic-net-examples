// Title: C# Aspose.Cells Example: Hide Gridlines in a Combo (Column‑Line) Chart
// Description: Demonstrates creating a workbook with sample data, adding a combo chart that mixes column and line series, and programmatically turning off both major and minor gridlines on the value and category axes to achieve a clean plot area. The result is saved as an XLSX file.
// Keywords: Aspose.Cells | C# | .NET Excel chart | combo chart | column line chart | hide chart gridlines | disable major gridlines | disable minor gridlines | plot area formatting | Excel visualization | global developers | US .NET community
// Common Searches: Aspose.Cells hide gridlines combo chart C# | remove major gridlines from Excel chart using Aspose.Cells | disable minor gridlines in column‑line chart .NET | turn off chart axes gridlines Aspose.Cells | C# code to hide plot area gridlines in Excel
// Developer Intent: Programmatically turn off all gridlines in the plot area of a combo chart.
// Use Cases: Generate dashboard‑style Excel reports with uncluttered charts. | Create presentation‑ready combo charts for sales or KPI data. | Automate report generation where background gridlines distract from trends. | Produce clean visual designs for international audiences (US, EU, APAC). | Integrate chart styling into CI pipelines for automated Excel output.
// AI Prompts: Write C# code using Aspose.Cells to hide major and minor gridlines in a combo chart. | Show how to toggle gridline visibility on chart axes with a boolean flag in Aspose.Cells. | Provide an example that changes gridline color after hiding them in an Aspose.Cells chart. | Explain how to hide gridlines for other chart types (pie, bar) with Aspose.Cells .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Demonstrates creating a workbook with sample data, adding a combo chart that mixes column and line series, and programmatically turning off both major and minor gridlines on the value and category axes to achieve a clean plot area. The result is saved as an XLSX file.
class HideChartGridlines
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a combo chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("ColumnSeries");
            sheet.Cells["C1"].PutValue("LineSeries");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["C4"].PutValue(35);

            // Add a combo chart (column + line)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // First series as column
            chart.NSeries.Add("B2:B4", true);
            // Second series as line to create a combo effect
            chart.NSeries.Add("C2:C4", true);
            chart.NSeries[1].Type = ChartType.Line;

            // Set category (X‑axis) data
            chart.NSeries.CategoryData = "A2:A4";

            // Hide all gridlines in the plot area
            chart.ValueAxis.MajorGridLines.IsVisible = false;
            chart.ValueAxis.MinorGridLines.IsVisible = false;
            chart.CategoryAxis.MajorGridLines.IsVisible = false;
            chart.CategoryAxis.MinorGridLines.IsVisible = false;

            // Save the workbook
            string outputPath = "ComboChart_NoGridlines.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

class Program
{
    static void Main(string[] args)
    {
        HideChartGridlines.Run();
    }
}
