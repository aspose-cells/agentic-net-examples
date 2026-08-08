// Title: Add a Leading Apostrophe with QuotePrefix StyleFlag in AspNet Cells for .NET (C#)
// Description: Demonstrates how to force Excel to treat a value as text by applying a Style with QuotePrefix = true via a StyleFlag. The example creates a workbook, writes a numeric string to A1, sets QuotePrefix only, and saves the file, resulting in a leading apostrophe that preserves the original text format.
// Keywords: Aspose.Cells QuotePrefix | StyleFlag QuotePrefix C# | add leading apostrophe Excel | force text cell Aspose | prevent numeric conversion Aspose.Cells | .NET Excel styling | SetStyle QuotePrefix example
// Common Searches: Aspose.Cells add leading apostrophe | QuotePrefix StyleFlag C# example | How to keep numeric strings as text in Excel using Aspose | SetStyle only QuotePrefix Aspose.Cells | Prevent Excel auto‑formatting with Aspose.Cells
// Developer Intent: Apply a QuotePrefix style to a cell using a StyleFlag so the value is stored with a leading apostrophe and treated as text.
// Use Cases: Store account numbers or IDs with leading zeros without losing formatting. | Export product or part codes that consist solely of digits but must remain strings. | Avoid scientific notation for large numeric strings when opening the workbook in Excel.
// AI Prompts: Generate C# code to apply QuotePrefix to an entire column using a StyleFlag in Aspose.Cells. | Show how to toggle QuotePrefix based on cell content (digits vs. text) with Aspose.Cells for .NET. | Explain the interaction between Style, StyleFlag, and SetStyle when only QuotePrefix is modified.

using System;
using Aspose.Cells;

namespace AsposeCellsQuotePrefixDemo
{
    // Demonstrates how to force Excel to treat a value as text by applying a Style with QuotePrefix = true via a StyleFlag. The example creates a workbook, writes a numeric string to A1, sets QuotePrefix only, and saves the file, resulting in a leading apostrophe that preserves the original text format.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Access a cell and put a value that should be treated as text
            Cell cell = worksheet.Cells["A1"];
            cell.PutValue("123456"); // Without apostrophe it would be a number

            // Create a style and enable QuotePrefix
            Style style = workbook.CreateStyle();
            style.QuotePrefix = true; // Indicates the value starts with a single quote

            // Create a StyleFlag and enable the QuotePrefix flag
            StyleFlag flag = new StyleFlag();
            flag.QuotePrefix = true; // Apply only the QuotePrefix setting

            // Apply the style to the cell using the flag
            cell.SetStyle(style, flag);

            // Save the workbook to verify the effect
            workbook.Save("QuotePrefixDemo.xlsx");
        }
    }
}
