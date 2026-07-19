// Title: Apply a Custom Font to All Chart Titles in an Aspose.Cells Workbook (C#)
// Description: Creates a workbook, adds a sample chart, then loops through every worksheet and chart to ensure the title is visible and sets its Font.Name to a branding font (e.g., Calibri). The workbook is saved with the updated chart title fonts.
// Keywords: Aspose.Cells chart title font | C# set chart title font | apply custom font to Excel chart titles | iterate worksheets charts Aspose | branding font for chart titles | Chart.Title.Font.Name Aspose.Cells | global chart title styling | Excel automation font family | consistent chart title appearance
// Common Searches: how to set the same font for all chart titles using Aspose.Cells C# | Aspose.Cells change chart title font family across workbook | C# code to apply branding font to Excel chart titles with Aspose | iterate all worksheets and charts to modify title font Aspose.Cells | set chart title font globally in a workbook
// Developer Intent: Programmatically assign a specific font family to every chart title in an Aspose.Cells workbook.
// Use Cases: Enforce corporate branding by applying the company’s standard font to all chart titles before distribution. | Maintain visual consistency across multiple sheets when generating reports automatically. | Create a template workbook where any newly added charts inherit the predefined title font without manual editing.
// AI Prompts: Generate C# code with Aspose.Cells that changes both the font size and color of all chart titles in an existing workbook. | Show how to apply a custom font to chart titles only when the title text contains a given keyword. | Provide a reusable method that accepts a font name, size, and color and applies it to every chart title across all worksheets.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds a sample chart, then loops through every worksheet and chart to ensure the title is visible and sets its Font.Name to a branding font (e.g., Calibri). The workbook is saved with the updated chart title fonts.
class ApplyCustomFontToChartTitles
{
    static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // -------------------------------------------------
        // Sample data and a chart (for demonstration only)
        // -------------------------------------------------
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);

        // Add a chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";
        chart.Title.Text = "Sample Chart";

        // -------------------------------------------------
        // Apply a custom font family to every chart title
        // -------------------------------------------------
        string brandingFont = "Calibri"; // replace with your desired font family

        foreach (Worksheet ws in workbook.Worksheets)
        {
            foreach (Chart ch in ws.Charts)
            {
                // Ensure the title is visible before applying font settings
                ch.Title.IsVisible = true;

                // Set the font family for the chart title
                ch.Title.Font.Name = brandingFont;
            }
        }

        // Save the workbook
        string outputPath = "WorkbookWithCustomChartTitleFont.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to {outputPath}");
    }
}
