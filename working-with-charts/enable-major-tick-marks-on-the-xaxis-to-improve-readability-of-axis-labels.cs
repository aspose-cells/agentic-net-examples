// Title: Aspose.Cells C# – Enable Major Tick Marks on the X‑Axis (Category Axis) of a Column Chart
// Description: Demonstrates how to create a workbook with sample data, add a column chart, and activate major tick marks on the X‑axis (category axis) using Aspose.Cells for .NET to make axis labels easier to read.
// Keywords: Aspose.Cells X‑axis tick marks C# | CategoryAxis MajorTickMark .NET | Excel chart tick mark formatting | Enable major tick marks Aspose.Cells | C# chart axis customization
// Common Searches: how to add major tick marks to chart X axis using Aspose.Cells | Aspose.Cells CategoryAxis MajorTickMark property example | C# set tick mark type for Excel chart axis | Aspose.Cells enable outside tick marks on category axis | chart axis formatting with Aspose.Cells .NET
// Developer Intent: Add visible major tick marks to the X‑axis of a generated chart to improve label clarity.
// Use Cases: Generate a column chart from data and enhance readability by showing outside tick marks on the category axis. | Apply a specific TickMarkType (Inside, Outside, Cross) to any chart axis in an automated Excel report. | Standardize axis appearance across multiple charts in a workbook before distribution.
// AI Prompts: Provide C# code with Aspose.Cells that sets major tick marks on the category axis of a line chart. | Show how to change the tick mark style to Inside for the value axis of a bar chart using Aspose.Cells. | Explain how to programmatically configure axis lines and tick marks for several charts in one workbook with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook with sample data, add a column chart, and activate major tick marks on the X‑axis (category axis) using Aspose.Cells for .NET to make axis labels easier to read.
class EnableMajorTickMarksXaxis
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the chart data source
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable major tick marks on the X‑axis (category axis)
        chart.CategoryAxis.MajorTickMark = TickMarkType.Outside;

        // Save the workbook to a file
        workbook.Save("EnableMajorTickMarksXaxis.xlsx");
    }
}
