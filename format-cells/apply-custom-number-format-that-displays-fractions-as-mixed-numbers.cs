// Title: C# – Apply Custom Mixed‑Number Format (# ??/??) in Aspose.Cells
// Description: Demonstrates how to create a workbook, insert a decimal value, define a style with the custom number format "# ??/??" to display mixed fractions (e.g., 2 ¾), apply the format using a StyleFlag, and save the file as MixedNumberFraction.xlsx.
// Keywords: Aspose.Cells | C# | custom number format | mixed number | fraction format | # ??/?? | StyleFlag | Excel export
// Common Searches: Aspose.Cells mixed number format C# | display fractions as mixed numbers in Excel using Aspose | custom number format # ??/?? Aspose.Cells | apply number format only with StyleFlag
// Developer Intent: Format cells so numeric values appear as mixed numbers (integer plus fraction) in an Excel file generated with Aspose.Cells.
// Use Cases: Present measurement data (e.g., 2 ¾ inches) in engineering reports. | Show recipe quantities such as 1 ½ cups in exported spreadsheets. | Render financial ratios with fractional components for clearer readability.
// AI Prompts: Write C# code that applies the "# ??/??" mixed‑number format to an entire column in Aspose.Cells. | Explain how StyleFlag can isolate number‑format changes while keeping other cell styles unchanged. | Provide a snippet that formats a range of decimal values as mixed fractions using Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, insert a decimal value, define a style with the custom number format "# ??/??" to display mixed fractions (e.g., 2 ¾), apply the format using a StyleFlag, and save the file as MixedNumberFraction.xlsx.
class MixedNumberFractionDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Set a numeric value that will be shown as a mixed fraction (e.g., 2.75 -> 2 3/4)
        worksheet.Cells["A1"].PutValue(2.75);

        // Create a style with a custom mixed‑number format
        // "# ??/??" displays the integer part and the fractional part as a mixed number
        Style mixedNumberStyle = workbook.CreateStyle();
        mixedNumberStyle.Custom = "# ??/??";

        // Use a StyleFlag to apply only the number format part of the style
        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true;

        // Apply the style to the target cell (or range)
        Aspose.Cells.Range range = worksheet.Cells.CreateRange("A1");
        range.ApplyStyle(mixedNumberStyle, flag);

        // Save the workbook
        workbook.Save("MixedNumberFraction.xlsx");
    }
}
