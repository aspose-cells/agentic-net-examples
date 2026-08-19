// Title: C# – Apply a Custom Font to All Chart Titles in an Aspose.Cells Workbook
// Description: Demonstrates how to iterate through every worksheet and chart in a workbook, make each chart title visible, and set a custom font family, size, and color for consistent branding using Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart title font | C# set chart title style | bulk update chart titles Aspose | custom font family Excel chart | .NET chart title formatting | branding Excel charts programmatically
// Common Searches: change font of all chart titles Aspose.Cells C# | set chart title font family in .NET workbook | apply branding font to Excel chart titles programmatically | make chart titles visible and style them Aspose.Cells | bulk modify chart title appearance in Excel file
// Developer Intent: Apply a specific font family, size, and color to every chart title in a workbook.
// Use Cases: Standardize chart titles with corporate branding before exporting reports. | Ensure consistent visual style for dashboards generated from templates. | Automate font updates across multiple worksheets in a shared workbook.
// AI Prompts: Generate C# code that loops through all worksheets in an Aspose.Cells workbook and sets each chart title to the font 'Helvetica', size 12, color red. | Provide a snippet that checks if a chart title is hidden, makes it visible, and then applies a custom font using Aspose.Cells for .NET. | Create a reusable method that accepts font name, size, and color parameters and applies them to every chart title in a given workbook.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to iterate through every worksheet and chart in a workbook, make each chart title visible, and set a custom font family, size, and color for consistent branding using Aspose.Cells for .NET.
class ApplyCustomChartTitleFont
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // -------------------------------------------------
        // Sample data and a chart to demonstrate the logic
        // -------------------------------------------------
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);

        // Add a chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";
        chart.Title.Text = "Sample Chart";

        // -------------------------------------------------
        // Apply a custom font family to every chart title
        // -------------------------------------------------
        string customFontFamily = "Calibri";

        foreach (Worksheet ws in workbook.Worksheets)
        {
            foreach (Chart ch in ws.Charts)
            {
                // Ensure the title is visible
                ch.Title.IsVisible = true;

                // Set the desired font family (and optional styling)
                ch.Title.Font.Name = customFontFamily;
                ch.Title.Font.Size = 14;
                ch.Title.Font.Color = Color.DarkBlue;
            }
        }

        // Save the workbook with the updated chart title fonts
        workbook.Save("CustomChartTitleFont.xlsx");
    }
}
