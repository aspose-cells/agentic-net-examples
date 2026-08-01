// Title: Apply Text Wrapping via CopyStyle in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a wrap‑enabled style, apply it to a source range, copy the style to another range with CopyStyle, insert long text, auto‑fit rows, and save the workbook, ensuring automatic text wrapping inheritance across cells.
// Keywords: Aspose.Cells CopyStyle | text wrap style .NET | C# Aspose.Cells wrap inheritance | AutoFitRows text wrapping | copy cell style Aspose | wrap enabled style Aspose.Cells | Excel export C# Aspose | range style copy Aspose
// Common Searches: Aspose.Cells copy style with text wrap | How to inherit IsTextWrapped property in Aspose.Cells | CopyStyle preserve wrap setting C# | AutoFitRows after applying wrap style Aspose | Apply text wrapping to multiple cells Aspose.Cells
// Developer Intent: Copy a wrap‑enabled style from one range to another so that long text automatically wraps in the destination cells.
// Use Cases: Define a single wrap style in a template and reuse it across many columns for consistent formatting. | Generate reports where description fields exceed column width, ensuring content stays readable without manual adjustments. | Apply a wrap style once, copy it to dynamic data ranges, and call AutoFitRows to adjust row heights automatically.
// AI Prompts: Provide C# code that creates a style with IsTextWrapped = true, applies it to range A1:A2, copies it to B1:B2 using CopyStyle, and saves the workbook with Aspose.Cells. | Explain why CopyStyle retains the IsTextWrapped property and how to combine it with AutoFitRows for correct row height. | Show an example that copies a wrap‑enabled style across multiple worksheets in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsWrapInheritanceDemo
{
    // Demonstrates how to create a wrap‑enabled style, apply it to a source range, copy the style to another range with CopyStyle, insert long text, auto‑fit rows, and save the workbook, ensuring automatic text wrapping inheritance across cells.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate source cells with long text that needs wrapping
                cells["A1"].PutValue("This is a very long text that should be wrapped automatically within the cell when the style is applied.");
                cells["A2"].PutValue("Another long text entry that demonstrates text wrapping inheritance across cells.");

                // Create a style with text wrapping enabled
                Style wrapStyle = workbook.CreateStyle();
                wrapStyle.IsTextWrapped = true;

                // Define a source range (A1:A2) and apply the wrapping style to it
                AsposeRange sourceRange = cells.CreateRange("A1:A2");
                sourceRange.SetStyle(wrapStyle);

                // Define a destination range (B1:B2) where the style will be inherited
                AsposeRange destinationRange = cells.CreateRange("B1:B2");
                // Copy the style from the source range to the destination range
                destinationRange.CopyStyle(sourceRange);

                // Populate destination cells with long text to see the inherited wrapping
                cells["B1"].PutValue("Destination cell with inherited wrap style. This text should also wrap automatically.");
                cells["B2"].PutValue("Second destination cell demonstrating inherited text wrapping functionality.");

                // Auto-fit rows to adjust height for wrapped text
                worksheet.AutoFitRows();

                // Determine output file path
                string outputPath = "WrapInheritanceDemo.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
