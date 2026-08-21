// Title: Apply a Custom Number Format in Aspose.Cells for .NET (C#) to Show Large IDs Without Scientific Notation
// Description: This example creates a workbook, writes three 13‑digit identifiers to cells A1‑A3, defines a custom style with the format "#,##0" to display numbers with thousand separators and without scientific notation, applies the style to the range using a StyleFlag that targets only the number format, and saves the file as LargeIdentifiersNumberFormat.xlsx.
// Keywords: Aspose.Cells | Aspose.Cells for .NET | C# Excel | custom number format | prevent scientific notation | large identifiers | plain numeric display | StyleFlag NumberFormat | Excel thousand separators | format cells programmatically
// Common Searches: Aspose.Cells C# custom number format example | how to stop scientific notation in Excel using Aspose.Cells | apply "#,##0" format to a range with Aspose.Cells .NET | set number format without affecting other styles Aspose.Cells | display large IDs as plain numbers in generated Excel file
// Developer Intent: The developer needs to format cells so that large identifier values are shown as full numbers rather than scientific notation.
// Use Cases: Exporting invoice numbers or order IDs that exceed 12 digits. | Preserving full SKU or barcode values in Excel reports. | Generating financial statements where account numbers must remain exact.
// AI Prompts: Write C# code that uses Aspose.Cells to apply the custom number format "#,##0" to cells A1‑A3 and prevents scientific notation. | Show how to use StyleFlag with NumberFormat = true to apply only the number format in Aspose.Cells. | Provide a complete Aspose.Cells .NET example that saves a workbook after formatting a column of large identifiers.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsNumberFormatDemo
{
    // This example creates a workbook, writes three 13‑digit identifiers to cells A1‑A3, defines a custom style with the format "#,##0" to display numbers with thousand separators and without scientific notation, applies the style to the range using a StyleFlag that targets only the number format, and saves the file as LargeIdentifiersNumberFormat.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate cells with large identifier values that would normally display in scientific notation
                sheet.Cells["A1"].PutValue(1234567890123);
                sheet.Cells["A2"].PutValue(9876543210987);
                sheet.Cells["A3"].PutValue(5555555555555);

                // Create a custom style that forces plain number display (no scientific notation)
                Style plainNumberStyle = workbook.CreateStyle();
                // Custom format "#,##0" adds thousand separators and prevents scientific notation
                plainNumberStyle.Custom = "#,##0";

                // Prepare a StyleFlag to apply only the number format part of the style
                StyleFlag flag = new StyleFlag
                {
                    NumberFormat = true
                };

                // Apply the style to the range containing the large identifiers
                AsposeRange range = sheet.Cells.CreateRange("A1", "A3");
                range.ApplyStyle(plainNumberStyle, flag);

                // Save the workbook to an Excel file
                workbook.Save("LargeIdentifiersNumberFormat.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
