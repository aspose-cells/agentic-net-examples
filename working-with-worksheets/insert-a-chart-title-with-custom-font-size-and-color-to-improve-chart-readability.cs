// Title: Set a custom font size and color for a column chart title using Aspose.Cells for .NET
// AI Prompts: Create a new Workbook, add a Column chart, and assign the title text "Quarterly Sales Report" with a 16‑point dark blue font using the Aspose.Cells C# API. | Adjust the Font.Size and Font.Color properties of an existing chart's Title object to customize its appearance before saving the workbook. | Generate sample sales data, bind it to the chart series, apply the custom title formatting, and save the file as an .xlsx workbook.
// Common Searches: aspnet set chart title font size and color Aspose.Cells example | how to change column chart title appearance with Aspose.Cells C# | customize Excel chart title color using Aspose.Cells for .NET | Aspose.Cells chart title formatting tutorial | set dark blue font for chart title in generated Excel file
// Tags: chart title font size Aspose.Cells | chart title color customization .NET | column chart title formatting C# | set chart title properties Aspose.Cells | custom chart title appearance Excel

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, inserts a column chart, sets the chart title to "Quarterly Sales Report" with a 16‑point dark blue font, adds sample sales data, binds the data to the chart series, and saves the workbook as ChartWithCustomTitle.xlsx.
class ChartTitleExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a column chart to the worksheet (position: row 5, column 0 to row 25, column 10)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the chart title text
            chart.Title.Text = "Quarterly Sales Report";

            // Customize the title font size
            chart.Title.Font.Size = 16;

            // Customize the title font color
            chart.Title.Font.Color = Color.DarkBlue;

            // (Optional) Add sample data for demonstration
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(12000);
            sheet.Cells["B3"].PutValue(15000);
            sheet.Cells["B4"].PutValue(13000);
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Save the workbook to a file
            workbook.Save("ChartWithCustomTitle.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
