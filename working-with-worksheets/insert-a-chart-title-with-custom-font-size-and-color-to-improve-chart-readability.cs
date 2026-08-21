// Title: Aspose.Cells .NET – Insert a Chart Title with Custom Font Size and Color
// Description: Creates a new workbook, adds sample sales data, inserts a column chart, makes the title visible, sets the title text to "Quarterly Sales", and formats the title with a 16‑point DarkBlue font before saving the file as an Excel workbook.
// Keywords: Aspose.Cells chart title font size | Aspose.Cells chart title color | C# set chart title Aspose.Cells | customize Excel chart title .NET | format chart title programmatically
// Common Searches: how to change chart title font size in Aspose.Cells | set chart title color using Aspose.Cells for .NET | add and style chart title in Excel with C# | Aspose.Cells custom chart title formatting
// Developer Intent: Add a visible chart title to a column chart and style it with a specific font size and color using Aspose.Cells for .NET.
// Use Cases: Produce sales dashboards where chart titles need consistent branding. | Automate generation of presentation‑ready Excel reports with styled titles. | Enforce corporate style guidelines for chart headings across financial workbooks.
// AI Prompts: Show me C# code to set the font family, size, and color of a chart title in Aspose.Cells. | How can I make a chart title bold, italic, and add a background color with Aspose.Cells? | Explain how to apply conditional formatting to chart titles based on data values in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a new workbook, adds sample sales data, inserts a column chart, makes the title visible, sets the title text to "Quarterly Sales", and formats the title with a 16‑point DarkBlue font before saving the file as an Excel workbook.
class InsertChartTitleCustom
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set the chart title text and make it visible
            chart.Title.Text = "Quarterly Sales";
            chart.Title.IsVisible = true;

            // Apply custom font size and color to the title
            chart.Title.Font.Size = 16;               // Custom font size
            chart.Title.Font.Color = Color.DarkBlue; // Custom font color

            // Save the workbook with the chart
            string outputPath = "ChartWithCustomTitle.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
