// Title: Add a "Monthly Revenue" chart title with Accent1 styling in Aspose.Cells for .NET
// Description: Generates a workbook, populates month and revenue rows, creates a column chart, makes the title visible, assigns the text "Monthly Revenue", and formats the title with the Accent1 theme color (RGB 0,112,192) and bold font before saving as ChartWithTitle.xlsx.
// Keywords: Aspose.Cells | C# chart title | Accent1 theme color | Excel chart formatting | column chart title styling | programmatic Excel chart | set chart title Aspose | apply predefined style | Excel automation C# | chart title color
// Common Searches: Aspose.Cells set chart title C# | Apply Accent1 color to Excel chart title | Change chart title font programmatically | C# column chart title formatting Aspose.Cells | Use theme colors for charts in .NET
// Developer Intent: Programmatically assign the text "Monthly Revenue" to a chart title and apply the Accent1 predefined color with bold styling.
// Use Cases: Produce monthly sales dashboards where every chart shares the corporate Accent1 title style. | Automate financial reporting worksheets that require consistent, branded chart headings. | Generate Excel files for client presentations with uniformly colored and bolded chart titles.
// AI Prompts: Show C# code using Aspose.Cells to set a chart title, make it visible, and apply the Accent1 theme color with bold text. | Explain how to retrieve any predefined Accent color from a workbook and use it for chart titles across different chart types. | Provide a step‑by‑step guide to format chart titles (font size, weight, color) using Aspose.Cells for .NET.

using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// Generates a workbook, populates month and revenue rows, creates a column chart, makes the title visible, assigns the text "Monthly Revenue", and formats the title with the Accent1 theme color (RGB 0,112,192) and bold font before saving as ChartWithTitle.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["B1"].PutValue("Revenue");
        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["B2"].PutValue(5000);
        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["B3"].PutValue(7000);
        sheet.Cells["A4"].PutValue("Mar");
        sheet.Cells["B4"].PutValue(6500);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Set the chart title
        chart.Title.Text = "Monthly Revenue";
        chart.Title.IsVisible = true;

        // Apply the predefined Accent1 style to the title (using the Accent1 color)
        chart.Title.Font.Color = Color.FromArgb(0, 112, 192); // Accent1 color
        chart.Title.Font.IsBold = true; // optional styling

        // Save the workbook
        workbook.Save("ChartWithTitle.xlsx");
    }
}
