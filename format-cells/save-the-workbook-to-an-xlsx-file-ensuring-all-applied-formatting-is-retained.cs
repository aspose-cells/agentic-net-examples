// Title: Save a workbook with cell styles to XLSX using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells save workbook to XLSX while keeping bold font and background color | How to export a formatted Excel file with Aspose.Cells .NET | Preserve cell formatting when saving an XLSX file using Aspose.Cells in C# | Aspose.Cells .NET write workbook to XLSX with style retention | Save Excel workbook with custom cell style using Aspose.Cells API
// Tags: Aspose.Cells save workbook to XLSX with styles | C# retain cell formatting in XLSX output | export formatted Excel file using Aspose.Cells .NET | preserve cell style when writing XLSX with Aspose | apply bold font and background color before saving workbook

using Aspose.Cells;
using System.Drawing;

// The example creates a new Workbook, applies a bold font and yellow solid background to cell A1, and then saves the file as "output.xlsx" using SaveFormat.Xlsx, ensuring all cell formatting is retained.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Example of applying formatting to demonstrate retention
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Put a value in cell A1
        Cell cell = cells["A1"];
        cell.PutValue("Formatted Text");

        // Create a style with bold font and yellow background
        Style style = cell.GetStyle();
        style.Font.IsBold = true;
        style.ForegroundColor = Color.Yellow;
        style.Pattern = BackgroundType.Solid;

        // Apply the style to the cell
        cell.SetStyle(style);

        // Save the workbook to an XLSX file, preserving all formatting
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
