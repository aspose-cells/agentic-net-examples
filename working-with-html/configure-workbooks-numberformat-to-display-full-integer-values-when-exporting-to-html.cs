// Title: Set integer‑only NumberFormat in AspNet Cells and export to HTML (C#)
// Description: Demonstrates how to create a style with the built‑in number format "0", apply it to a worksheet range using StyleFlag, and save the workbook as HTML so all integer values appear fully without decimals or scientific notation.
// Keywords: Aspose.Cells integer format HTML | C# number format 0 Aspose.Cells | export Excel to HTML without decimals | StyleFlag NumberFormat Aspose.Cells | HtmlSaveOptions integer display
// Common Searches: Aspose.Cells show full integers in HTML export | C# set number format to 0 before saving as HTML | apply integer only style to range Aspose.Cells | prevent scientific notation in HTML output Aspose.Cells
// Developer Intent: Apply a number format that forces integer display to cells and preserve that formatting when converting the workbook to HTML.
// Use Cases: Create a Style with Number = 1 (pattern "0") and apply it to the used range via StyleFlag.NumberFormat = true. | Export a workbook containing large or negative integers to HTML without rounding or scientific notation. | Maintain existing cell styles while overriding only the numeric format for HTML output.
// AI Prompts: Generate C# code that forces integer display for all cells in an Aspose.Cells workbook before saving as HTML. | Show how to use StyleFlag to apply only the NumberFormat to a range and then export to HTML. | Explain the minimal HtmlSaveOptions configuration needed when integer‑only formatting is already set.

using System;
using Aspose.Cells;

namespace AsposeCellsNumberFormatHtml
{
    // Demonstrates how to create a style with the built‑in number format "0", apply it to a worksheet range using StyleFlag, and save the workbook as HTML so all integer values appear fully without decimals or scientific notation.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some integer values, including a large integer
                sheet.Cells["A1"].PutValue(123);
                sheet.Cells["A2"].PutValue(4567890123);
                sheet.Cells["A3"].PutValue(0);
                sheet.Cells["A4"].PutValue(-987654321);

                // Create a style that forces integer display (no decimal places)
                // Built‑in number format 1 corresponds to the pattern "0"
                Style intStyle = workbook.CreateStyle();
                intStyle.Number = 1; // "0" format

                // Apply the style to the used range
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
                StyleFlag flag = new StyleFlag();
                flag.NumberFormat = true; // Apply only the number format
                usedRange.ApplyStyle(intStyle, flag);

                // Configure HTML save options (default options are sufficient)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

                // Save the workbook as HTML; integers will be shown fully
                workbook.Save("IntegerValues.html", htmlOptions);

                Console.WriteLine("HTML file generated with full integer display.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
