// Title: Apply a custom number format to display negative percentages in red with a minus sign using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# custom number format red negative percentage | set cell style to show negative percentages in red in Excel with Aspose.Cells | C# format negative percent values with minus sign and red color using Aspose.Cells | how to use a custom format string for red negative percentages in Aspose.Cells .NET | apply conditional color formatting to negative percentages in an Aspose.Cells workbook
// Tags: Aspose.Cells custom format string | red negative percent cell styling | C# Excel cell style application | set cell style Aspose.Cells workbook | export workbook to .xlsx Aspose.Cells

using Aspose.Cells;

// The example creates a new workbook, writes a negative decimal (-0.1234) to cell A1, defines a custom number format "0.00%;[Red]-0.00%" that shows positive percentages normally and negative percentages in red with a minus sign, applies this style to the cell, and saves the workbook as NegativePercentages.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();

        // Access the first worksheet
        var sheet = workbook.Worksheets[0];

        // Insert a negative percentage value into cell A1
        var cell = sheet.Cells["A1"];
        cell.PutValue(-0.1234); // Represents -12.34%

        // Define a custom number format: positive percentages normal, negative percentages red with minus sign
        var style = workbook.CreateStyle();
        style.Custom = "0.00%;[Red]-0.00%";

        // Apply the custom style to the cell
        cell.SetStyle(style);

        // Save the workbook to a file
        workbook.Save("NegativePercentages.xlsx");
    }
}
