// Title: Create a column chart and set its series colors to the workbook’s Theme Accent5 using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set column chart series color to workbook theme Accent5 | How to retrieve ThemeColorType.Accent5 and apply to chart series in Aspose.Cells | Apply workbook theme colors to Excel chart series with Aspose.Cells .NET | C# example using GetThemeColor to color chart series in an XLSX file | Aspose.Cells column chart series foreground and border color from theme
// Tags: set chart series color using workbook theme Aspose.Cells | retrieve theme accent5 color Aspose.Cells C# | column chart series formatting Aspose.Cells | apply theme colors to Excel chart series .NET | Aspose.Cells chart series foreground color

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, adds sample data, inserts a column chart, retrieves the workbook's Accent5 theme color via GetThemeColor, and assigns that color to each series' foreground and border before saving the file as ChartWithAccent5.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["C1"].PutValue("Profit");

            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            sheet.Cells["C2"].PutValue(30);
            sheet.Cells["C3"].PutValue(45);
            sheet.Cells["C4"].PutValue(60);

            // Add a column chart to the worksheet
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIdx];

            // Add series for Sales and Profit
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.Add("C2:C4", true);

            // Retrieve the workbook's theme Accent5 color
            Color accent5 = workbook.GetThemeColor(ThemeColorType.Accent5);

            // Apply Accent5 color to each series
            foreach (Series series in chart.NSeries)
            {
                // Fill color for columns/bars
                series.Area.ForegroundColor = accent5;
                // Optional: set border color if needed
                series.Border.Color = accent5;
            }

            // Save the workbook with the chart
            workbook.Save("ChartWithAccent5.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
