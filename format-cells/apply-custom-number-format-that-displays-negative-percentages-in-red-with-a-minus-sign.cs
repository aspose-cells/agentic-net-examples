// Title: How to format negative percentages in red with a minus sign using Aspose.Cells for .NET
// AI Prompts: Create a C# style in Aspose.Cells using the format pattern 0.00%;[Red]-0.00% and assign it to percentage cells. | Generate an Excel file where positive percentages appear normally while negative percentages are shown in red with a leading minus sign using Aspose.Cells. | Define a cell style that formats percentages with two decimals and red negative values, then save the workbook as FormattedPercentages.xlsx.
// Common Searches: Aspose.Cells C# format negative percentages in red | how to show red minus sign for negative percentages in Excel with Aspose.Cells | apply color to negative percentage values using Aspose.Cells API | custom format string for percentages Aspose.Cells .NET | save workbook with styled percentage cells Aspose.Cells
// Tags: percentage cell style with red negative values Aspose.Cells | custom number format pattern 0.00%;[Red]-0.00% Aspose | apply style to cells C# Aspose.Cells | xlsx output with formatted percentages Aspose.Cells | two-decimal percentage formatting .NET

using Aspose.Cells;

// Demonstrates creating a workbook, inserting positive and negative percentage values, defining a custom number format "0.00%;[Red]-0.00%" to display negative percentages in red with a minus sign, applying the style to cells A1 and A2, and saving the file as FormattedPercentages.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();

        // Access the first worksheet
        var sheet = workbook.Worksheets[0];

        // Insert sample values: a positive and a negative percentage
        sheet.Cells["A1"].PutValue(0.1234);   // 12.34%
        sheet.Cells["A2"].PutValue(-0.0567); // -5.67%

        // Custom number format:
        //   Positive percentages: normal display
        //   Negative percentages: red color with a minus sign
        string customFormat = "0.00%;[Red]-0.00%";

        // Create a style with the custom format
        Style style = workbook.CreateStyle();
        style.Custom = customFormat;

        // Apply the style to the cells
        sheet.Cells["A1"].SetStyle(style);
        sheet.Cells["A2"].SetStyle(style);

        // Save the workbook
        workbook.Save("FormattedPercentages.xlsx");
    }
}
