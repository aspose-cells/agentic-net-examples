// Title: Export Large Numbers to HTML Without Scientific Notation Using Aspose.Cells for .NET
// Description: Demonstrates how to insert a very large numeric value into a workbook, set the SignificantDigitsType to G17, and save the sheet as HTML so the number appears in plain decimal format instead of scientific notation.
// Keywords: Aspose.Cells | C# | HTML export | large numbers | prevent scientific notation | SignificantDigitsType | G17 format | Excel to HTML | full‑precision display | US developers | European .NET community
// Common Searches: Aspose.Cells export large number to HTML without exponential notation | C# prevent scientific notation in HTML output using Aspose.Cells | Set SignificantDigitsType G17 for HTML export in Aspose.Cells | How to keep big integers readable in HTML saved from Excel | Aspose.Cells number format options for web reports
// Developer Intent: The developer needs to generate an HTML file from a workbook while ensuring that very large numeric values are rendered as full decimal strings rather than scientific (exponential) notation.
// Use Cases: Financial dashboards where account numbers exceed 15 digits and must stay legible in a web view. | Web‑based invoices that contain product or serial codes larger than typical integer limits. | Scientific data portals that publish high‑precision identifiers without converting them to exponential form.
// AI Prompts: Write C# code with Aspose.Cells that saves a worksheet to HTML, displaying a 20‑digit number in plain decimal using SignificantDigitsType.G17. | Explain the impact of SignificantDigitsType on number formatting in Aspose.Cells and recommend the best setting for HTML exports of large values. | Show how to programmatically verify that the generated HTML contains the exact numeric string and not scientific notation.

using System;
using Aspose.Cells;

// Demonstrates how to insert a very large numeric value into a workbook, set the SignificantDigitsType to G17, and save the sheet as HTML so the number appears in plain decimal format instead of scientific notation.
class PreventExponentialHtmlExport
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Insert a large numeric value that would normally appear in exponential notation
        sheet.Cells["A1"].PutValue(12345678901234567890.0);

        // Adjust number format handling: use G17 format to avoid exponential notation
        // This sets the global significant digits type for the workbook
        workbook.Settings.SignificantDigitsType = SignificantDigitsType.G17;

        // Prepare HTML save options (no special options needed beyond defaults)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Save the workbook as HTML (lifecycle: save)
        workbook.Save("LargeNumber.html", htmlOptions);
    }
}
