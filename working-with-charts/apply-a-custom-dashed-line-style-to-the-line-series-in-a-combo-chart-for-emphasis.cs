// Title: C# – Apply a Red Dash‑Dot Line Style to the Line Series of an Aspose.Cells Combo Chart
// Description: Creates a workbook, adds month, sales, and target data, builds a combo chart with a column series for sales and a line series for target, converts the second series to a line type, and formats that line with a red dash‑dot border of medium thickness before saving the file.
// Keywords: Aspose.Cells | combo chart | line series formatting | dash dot border | C# chart styling | Excel automation | MsoLineDashStyle | WeightType | chart series border color | custom chart line style
// Common Searches: Aspose.Cells set dash‑dot line style for chart series C# | How to change line series border in an Aspose.Cells combo chart | C# Aspose.Cells custom line formatting in Excel charts | Apply red dash‑dot border to line series using Aspose.Cells | Combo chart line series styling Aspose.Cells .NET
// Developer Intent: Generate a combo chart with column and line series and apply a red dash‑dot border of medium weight to the line series using Aspose.Cells for .NET.
// Use Cases: Emphasize a target trend line in sales dashboards with a distinctive dash‑dot style. | Visually separate multiple data series in a single chart for clearer analysis. | Enforce corporate branding by programmatically applying specific line colors and dash patterns to Excel charts.
// AI Prompts: Write C# code with Aspose.Cells that creates a combo chart where the line series uses a blue dashed line with a thick border. | List all available dash styles in Aspose.Cells and show how to change the border dash type, color, and weight of a chart series. | Provide step‑by‑step instructions to modify an existing Aspose.Cells chart so the line series adopts a custom dash pattern.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

// Creates a workbook, adds month, sales, and target data, builds a combo chart with a column series for sales and a line series for target, converts the second series to a line type, and formats that line with a red dash‑dot border of medium thickness before saving the file.
class ComboChartCustomDash
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the combo chart
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["C1"].PutValue("Target");

        string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
        double[] sales = { 120, 150, 170, 130, 160 };
        double[] target = { 100, 140, 160, 120, 150 };

        for (int i = 0; i < months.Length; i++)
        {
            sheet.Cells[i + 1, 0].PutValue(months[i]);   // Column A: Month
            sheet.Cells[i + 1, 1].PutValue(sales[i]);   // Column B: Sales (column series)
            sheet.Cells[i + 1, 2].PutValue(target[i]);  // Column C: Target (line series)
        }

        // Add a combo chart (default type Column) to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Add the column series (Sales)
        chart.NSeries.Add("B2:B6", true);
        // Add the line series (Target)
        chart.NSeries.Add("C2:C6", true);
        // Set category (X‑axis) data
        chart.NSeries.CategoryData = "A2:A6";

        // Convert the second series to a line type
        chart.NSeries[1].Type = ChartType.Line;

        // Apply a custom dashed line style to the line series for emphasis
        chart.NSeries[1].Border.DashType = MsoLineDashStyle.DashDot; // dash‑dot pattern
        chart.NSeries[1].Border.Color = Color.Red;                  // line color
        chart.NSeries[1].Border.Weight = WeightType.MediumLine;    // line thickness

        // Save the workbook with the customized combo chart
        workbook.Save("ComboChartCustomDash.xlsx");
    }
}
