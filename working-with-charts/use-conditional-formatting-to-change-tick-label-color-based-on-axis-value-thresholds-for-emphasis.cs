// Title: Conditionally format Excel chart axis tick labels by value thresholds with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# change chart axis tick label color based on value | example of conditional formatting for category axis labels in an Excel chart using Aspose.Cells | how to set red font for axis labels greater than 50 with Aspose.Cells | color Excel chart axis labels green when they are less than 20 using Aspose.Cells C# | apply conditional formatting to chart axis labels programmatically with Aspose.Cells
// Tags: Aspose.Cells conditional formatting axis labels | chart tick label color thresholds Aspose.Cells | C# set axis label font color based on cell value | Excel chart category axis conditional formatting .NET | apply conditional formatting to chart axis using Aspose.Cells | conditional formatting range for chart labels Aspose

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills columns A and B with categories and values, adds a column chart, defines a conditional formatting range on cells A2:A5, and applies two rules: axis labels greater than 50 are colored red and labels less than 20 are colored green. The workbook is saved as ConditionalAxisLabelFormatting.xlsx.
class ConditionalAxisLabelFormatting
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue(10);
            sheet.Cells["A3"].PutValue(30);
            sheet.Cells["A4"].PutValue(55);
            sheet.Cells["A5"].PutValue(70);
            sheet.Cells["B2"].PutValue(15);
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["B4"].PutValue(60);
            sheet.Cells["B5"].PutValue(80);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (values from column B)
            chart.NSeries.Add("B2:B5", true);
            // Set the category axis labels (from column A)
            chart.NSeries.CategoryData = "A2:A5";

            // ------------------------------------------------------------
            // Apply conditional formatting to the axis (category) labels
            // ------------------------------------------------------------

            // Define the range that contains the axis labels (A2:A5)
            CellArea labelRange = CellArea.CreateCellArea("A2", "A5");

            // Add a new conditional formatting collection
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection labelConditions = sheet.ConditionalFormattings[cfIndex];
            labelConditions.AddArea(labelRange);

            // Condition 1: If label value > 50, set font color to Red
            int condIdx1 = labelConditions.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "50",
                string.Empty);
            FormatCondition condGreaterThan50 = labelConditions[condIdx1];
            condGreaterThan50.Style.Font.Color = Color.Red;

            // Condition 2: If label value < 20, set font color to Green
            int condIdx2 = labelConditions.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.LessThan,
                "20",
                string.Empty);
            FormatCondition condLessThan20 = labelConditions[condIdx2];
            condLessThan20.Style.Font.Color = Color.Green;

            // Save the workbook
            workbook.Save("ConditionalAxisLabelFormatting.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
