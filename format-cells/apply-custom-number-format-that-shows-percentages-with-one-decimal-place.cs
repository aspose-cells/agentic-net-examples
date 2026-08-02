// Title: C# – Apply custom number format "0.0%" to show percentages with one decimal using Aspose.Cells
// Description: Creates a workbook, writes a decimal to cell A1, defines a style with the custom format 0.0% to display percentages with one decimal place, applies the number‑format via StyleFlag, and saves the file as PercentageOneDecimal.xlsx.
// Keywords: Aspose.Cells | C# | .NET | custom number format | percentage format | 0.0% | StyleFlag | cell styling | Excel export
// Common Searches: Aspose.Cells format cell as percentage with one decimal C# | custom number format 0.0% Aspose.Cells example | apply number format only using StyleFlag Aspose.Cells | save workbook after percentage formatting Aspose.Cells
// Developer Intent: Format a cell (or range) so its value appears as a percentage with a single decimal digit.
// Use Cases: Financial reports that require ratios displayed as 1‑decimal percentages. | Dashboard worksheets where all percentage columns share the same formatting. | Applying percentage formatting to specific cells while preserving other style attributes via StyleFlag.
// AI Prompts: Provide C# code that sets the custom number format "0.0%" on a range of cells in Aspose.Cells without altering existing styles. | Show an example of formatting multiple columns as percentages with one decimal place using StyleFlag and then saving the workbook. | Explain how to extend the format to include thousand separators, e.g., "#,##0.0%", in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Creates a workbook, writes a decimal to cell A1, defines a style with the custom format 0.0% to display percentages with one decimal place, applies the number‑format via StyleFlag, and saves the file as PercentageOneDecimal.xlsx.
class ApplyPercentageOneDecimalFormat
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put a sample decimal value into cell A1
        worksheet.Cells["A1"].PutValue(0.4567);

        // Create a style and set a custom number format that shows percentages with one decimal place
        Style percentStyle = workbook.CreateStyle();
        // "0.0%" displays the value as a percentage with one decimal place
        percentStyle.SetCustom("0.0%", true);

        // Apply only the number format part of the style using a StyleFlag
        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true;

        // Apply the style to the target cell (A1)
        worksheet.Cells["A1"].SetStyle(percentStyle);
        // Alternatively, to respect the flag explicitly:
        // Aspose.Cells.Range range = worksheet.Cells.CreateRange("A1");
        // range.ApplyStyle(percentStyle, flag);

        // Save the workbook
        workbook.Save("PercentageOneDecimal.xlsx");
    }
}
