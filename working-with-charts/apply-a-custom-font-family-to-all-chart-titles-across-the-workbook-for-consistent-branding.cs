// Title: Set a custom font family for all chart titles in an Aspose.Cells workbook (C#)
// Description: Creates a workbook, adds sample data and a chart, then loops through every worksheet and chart to make the title visible and assign a specified font (e.g., Calibri) to each chart title before saving the file.
// Keywords: Aspose.Cells chart title font C# | global chart title formatting Aspose | set chart title font family .NET | iterate worksheets charts Aspose.Cells | apply corporate font to chart titles
// Common Searches: How to change the font of all chart titles in Aspose.Cells using C# | Aspose.Cells set chart title font for multiple sheets | C# code to apply a custom font to every chart title in a workbook | Batch update chart title fonts with Aspose.Cells
// Developer Intent: Apply a specific font family to every chart title across all worksheets in a workbook.
// Use Cases: Enforce corporate branding by using the company’s standard font on all chart titles. | Prepare automated reports where chart titles need a consistent, readable typeface. | Migrate legacy workbooks to a new visual style by updating chart title fonts in bulk.
// AI Prompts: Generate C# code with Aspose.Cells that iterates through all worksheets and charts to set Title.IsVisible = true and Font.Name to a given font. | Show how to also change the size and color of each chart title while keeping the custom font family. | Explain step‑by‑step how to globally apply a custom font to chart titles in an Aspose.Cells workbook and save the result.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartTitleFontDemo
{
    // Creates a workbook, adds sample data and a chart, then loops through every worksheet and chart to make the title visible and assign a specified font (e.g., Calibri) to each chart title before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Sample data and chart creation (for demonstration)
            // -------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set an initial title (optional)
            chart.Title.Text = "Sample Chart";
            chart.Title.IsVisible = true;

            // -------------------------------------------------
            // Apply custom font family to all chart titles
            // -------------------------------------------------
            string customFontName = "Calibri"; // Desired font family

            // Iterate through every worksheet in the workbook
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Iterate through every chart in the current worksheet
                foreach (Chart ch in ws.Charts)
                {
                    // Ensure the title is visible before applying font settings
                    ch.Title.IsVisible = true;

                    // Apply the custom font family to the chart title
                    ch.Title.Font.Name = customFontName;

                    // (Optional) Additional font styling can be set here, e.g. size or color
                    // ch.Title.Font.Size = 14;
                    // ch.Title.Font.Color = Color.DarkBlue;
                }
            }

            // Save the workbook with the updated chart title fonts
            workbook.Save("Workbook_With_CustomChartTitleFont.xlsx");
        }
    }
}
