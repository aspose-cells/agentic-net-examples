// Title: Apply a custom number format in Aspose.Cells for .NET to show negative percentages in red with a minus sign (C#)
// Common Searches: Aspose.Cells C# custom format red negative percentage | How to display negative percent values in red using Aspose.Cells for .NET | C# Excel custom number format 0.00% red negative cells Aspose | Set style for negative percentages with minus sign in Aspose.Cells workbook | Formatting percentage cells with color conditions in Aspose.Cells C#
// Tags: custom number format negative percentages Aspose.Cells | red color formatting for negative percent cells C# | Aspose.Cells style custom percentage format | Excel workbook percentage styling Aspose.NET | C# apply custom format to cells Aspose

using Aspose.Cells;

// The example creates a workbook, writes positive and negative decimal values, defines a custom number format "0.00%;[Red]-0.00%" that renders negative percentages in red with a leading minus sign, applies this style to the target cells, and saves the file as FormattedPercentages.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data: positive and negative percentages
        sheet.Cells["A1"].PutValue(0.1234);   // 12.34%
        sheet.Cells["A2"].PutValue(-0.0567); // -5.67%

        // Custom number format:
        // Positive percentages: normal (e.g., 12.34%)
        // Negative percentages: red color with minus sign (e.g., -5.67%)
        string customNumberFormat = "0.00%;[Red]-0.00%";

        // Create a style and assign the custom format
        Style percentStyle = workbook.CreateStyle();
        percentStyle.Custom = customNumberFormat;

        // Apply the style to the target cells
        sheet.Cells["A1"].SetStyle(percentStyle);
        sheet.Cells["A2"].SetStyle(percentStyle);

        // Save the workbook (lifecycle rule)
        workbook.Save("FormattedPercentages.xlsx");
    }
}
