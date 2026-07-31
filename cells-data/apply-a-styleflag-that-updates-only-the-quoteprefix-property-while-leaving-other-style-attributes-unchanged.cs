// Title: Set QuotePrefix Only with StyleFlag in Aspise.Cells for .NET
// Description: Demonstrates how to enable the QuotePrefix flag on a single cell (or range) using a StyleFlag so that only the QuotePrefix attribute is updated while all other formatting remains untouched. The example creates a workbook, writes a numeric‑looking string, applies the style, and saves the file.
// Keywords: Aspose.Cells QuotePrefix | StyleFlag QuotePrefix only | C# SetStyle QuotePrefix | Excel treat value as text | Preserve cell formatting Aspose
// Common Searches: how to set QuotePrefix with StyleFlag Aspose.Cells | apply only QuotePrefix style in C# Excel library | preserve existing cell format while enabling QuotePrefix | Aspose.Cells SetStyle flag examples | prevent Excel auto‑formatting numeric strings
// Developer Intent: Enable the QuotePrefix property on targeted cells without altering any other style settings.
// Use Cases: Keep leading zeros in IDs or product codes when exporting data to Excel. | Prevent Excel from auto‑converting numeric strings to numbers in generated reports. | Update only the QuotePrefix of a range while retaining existing fonts, colors, and borders.
// AI Prompts: Show C# code that uses StyleFlag to set only QuotePrefix on a cell with Aspose.Cells. | Explain how StyleFlag works with SetStyle to modify a single style attribute in Aspose.Cells. | Generate an example that toggles QuotePrefix for multiple cells without resetting other formatting.

using Aspose.Cells;
using System;

// Demonstrates how to enable the QuotePrefix flag on a single cell (or range) using a StyleFlag so that only the QuotePrefix attribute is updated while all other formatting remains untouched. The example creates a workbook, writes a numeric‑looking string, applies the style, and saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Access cell A1 and put a value that looks like a number
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("12345");

        // Create a style and enable the QuotePrefix property
        Style style = workbook.CreateStyle();
        style.QuotePrefix = true;   // the value will be treated as text

        // Create a StyleFlag and set only the QuotePrefix flag to true
        StyleFlag flag = new StyleFlag();
        flag.QuotePrefix = true;    // ensures only QuotePrefix is applied

        // Apply the style to the cell using the flag
        cell.SetStyle(style, flag);

        // Save the workbook
        workbook.Save("QuotePrefixOnly.xlsx");
    }
}
