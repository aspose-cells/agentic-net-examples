// Title: Format a cell as a percentage with one decimal place using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# format cell as percentage with one decimal place | How to apply custom number format 0.0% to a cell in Aspose.Cells | Set custom percentage style for a worksheet cell using Aspose.Cells .NET | Save an Excel file after applying one-decimal-place percentage format with Aspose.Cells
// Tags: custom percentage number format Aspose.Cells | apply cell style custom format .NET | percentage one decimal place formatting C# | Aspose.Cells workbook save with custom format | cell style percentage format example

using Aspose.Cells;

// The example creates a new workbook, writes a decimal value to cell A1, defines a custom number format "0.0%" to display the value as a percentage with one decimal place (e.g., 12.3%), applies the style to the cell, and saves the workbook as PercentageOneDecimal.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Put a sample decimal value into cell A1
        Cell cell = sheet.Cells["A1"];
        cell.PutValue(0.1234); // Represents 12.34%

        // Create a style with a custom number format for percentages with one decimal place
        Style style = cell.GetStyle();
        style.Custom = "0.0%"; // Displays as 12.3%
        cell.SetStyle(style);

        // Save the workbook to a file
        workbook.Save("PercentageOneDecimal.xlsx");
    }
}
