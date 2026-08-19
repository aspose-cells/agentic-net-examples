// Title: Aspose.Cells C# – Apply Conditional Formatting and Dynamically Color Chart Axis Tick Labels
// Description: Shows how to create a workbook, fill categories and values, add a column chart, apply a red‑yellow‑green three‑color conditional format to the data cells, link the value‑axis tick‑label number format to the worksheet, and change the tick‑label font to red when any value exceeds a specified threshold (e.g., 40). The result is saved as an XLSX file.
// Keywords: Aspose.Cells | C# | conditional formatting | chart axis styling | tick label color | value threshold | three color scale | Excel column chart | dynamic label formatting | SaveFormat.Xlsx
// Common Searches: Aspose.Cells change chart axis label color based on data | apply three color conditional format to Excel range using C# | link chart tick label number format to worksheet cells Aspose | set dynamic font color for axis labels in Aspose.Cells chart | C# example conditional formatting with Aspose.Cells
// Developer Intent: Programmatically style a chart’s axis labels according to data values with Aspose.Cells for .NET.
// Use Cases: Highlight extreme values in a report by coloring axis labels when a threshold is crossed. | Create dashboards where cell conditional formats automatically drive chart label appearance. | Generate Excel files that keep axis label formatting in sync with underlying data updates.
// AI Prompts: Modify the sample to use a four‑color scale (red‑orange‑yellow‑green) and change the tick‑label color when values exceed 70. | Write a reusable C# method that accepts a workbook, a data range, and a numeric limit, then applies a conditional color scale and updates the chart’s value‑axis tick‑label font accordingly.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to create a workbook, fill categories and values, add a column chart, apply a red‑yellow‑green three‑color conditional format to the data cells, link the value‑axis tick‑label number format to the worksheet, and change the tick‑label font to red when any value exceeds a specified threshold (e.g., 40). The result is saved as an XLSX file.
class ConditionalTickLabelColorDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data: categories in column A, numeric values in column B
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        string[] categories = { "Low", "Medium", "High", "Very High" };
        double[] values = { 5, 15, 30, 60 };
        for (int i = 0; i < categories.Length; i++)
        {
            sheet.Cells[i + 1, 0].PutValue(categories[i]); // A column (category)
            sheet.Cells[i + 1, 1].PutValue(values[i]);    // B column (value)
        }

        // Add a column chart that uses the data range
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B5", true);          // Values
        chart.NSeries.CategoryData = "A2:A5";      // Categories

        // ------------------------------------------------------------
        // Apply conditional formatting to the value cells (B2:B5)
        // ------------------------------------------------------------
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

        // Define the range that will receive the conditional formatting
        CellArea ca = new CellArea { StartRow = 1, EndRow = 4, StartColumn = 1, EndColumn = 1 };
        fcs.AddArea(ca);

        // Create a 3‑color scale: low = red, midpoint = yellow, high = green
        int condIdx = fcs.AddCondition(FormatConditionType.ColorScale);
        FormatCondition cond = fcs[condIdx];
        cond.ColorScale.Is3ColorScale = true;

        cond.ColorScale.MinCfvo.Type = FormatConditionValueType.Min;
        cond.ColorScale.MinColor = Color.Red;

        cond.ColorScale.MidCfvo.Type = FormatConditionValueType.Percentile;
        cond.ColorScale.MidCfvo.Value = 50; // 50th percentile (median)
        cond.ColorScale.MidColor = Color.Yellow;

        cond.ColorScale.MaxCfvo.Type = FormatConditionValueType.Max;
        cond.ColorScale.MaxColor = Color.Green;

        // ------------------------------------------------------------
        // Link the number format of the value axis tick labels to the cells
        // ------------------------------------------------------------
        chart.ValueAxis.TickLabels.NumberFormatLinked = true;
        chart.ValueAxis.TickLabels.NumberFormat = "0";

        // ------------------------------------------------------------
        // Change the overall tick‑label font color based on a threshold
        // (e.g., if any value exceeds 40, make the labels red)
        // ------------------------------------------------------------
        double maxVal = 0;
        foreach (double v in values)
        {
            if (v > maxVal) maxVal = v;
        }
        chart.ValueAxis.TickLabels.Font.Color = maxVal > 40 ? Color.Red : Color.Black;

        // Save the workbook
        workbook.Save("ConditionalTickLabelColorDemo.xlsx", SaveFormat.Xlsx);
    }
}
