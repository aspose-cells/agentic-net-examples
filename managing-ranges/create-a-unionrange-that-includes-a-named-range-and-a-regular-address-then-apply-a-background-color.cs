// Title: Aspose.Cells .NET – Build a UnionRange from a named range and a cell address and apply a background color
// Description: Demonstrates how to create a named range (MyRange = A1:A3), combine it with a regular address (C5:D6) into a UnionRange, and apply a solid LightGreen background using C# and Aspose.Cells. The workbook is saved as UnionRangeNamedDemo.xlsx.
// Keywords: Aspose.Cells UnionRange C# | named range Aspose.Cells | apply background color UnionRange | CreateUnionRange named range | .NET spreadsheet formatting | UnionRange style flag
// Common Searches: Aspose.Cells combine named range with cell range | UnionRange background color C# | How to use CreateUnionRange in Aspose.Cells .NET | Apply solid fill to multiple ranges Aspose.Cells
// Developer Intent: Generate a UnionRange that merges a named range and a standard address, then format the whole range with a solid fill.
// Use Cases: Highlight a header defined by a named range together with adjacent summary cells. | Apply uniform styling to a dynamic named range and fixed template cells in a report. | Group related sections of a worksheet by giving them a shared background color.
// AI Prompts: Provide C# code that creates a UnionRange from the named range MyRange and the address C5:D6, then sets a LightGreen background using Aspose.Cells. | Show an example of defining a named range, merging it with another range via CreateUnionRange, and applying a solid fill with StyleFlag.All in Aspose.Cells for .NET. | Explain the steps to format all parts of a UnionRange that includes both a named range and a regular cell block.

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsUnionRangeDemo
{
    // Demonstrates how to create a named range (MyRange = A1:A3), combine it with a regular address (C5:D6) into a UnionRange, and apply a solid LightGreen background using C# and Aspose.Cells. The workbook is saved as UnionRangeNamedDemo.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // 1. Create a named range "MyRange" that covers cells A1:A3
                // ------------------------------------------------------------
                AsposeRange rangeA = worksheet.Cells.CreateRange("A1:A3");
                rangeA.Name = "MyRange";

                // Populate the named range with sample data (optional)
                rangeA[0, 0].PutValue("Item 1");
                rangeA[1, 0].PutValue("Item 2");
                rangeA[2, 0].PutValue("Item 3");

                // ------------------------------------------------------------
                // 2. Create a UnionRange that includes the named range and a
                //    regular address (C5:D6)
                // ------------------------------------------------------------
                // The address string can contain both a named range and normal
                // cell addresses separated by commas.
                UnionRange unionRange = workbook.Worksheets.CreateUnionRange("MyRange,C5:D6", 0);

                // Populate the regular address part with sample data (optional)
                worksheet.Cells["C5"].PutValue("Data C5");
                worksheet.Cells["D6"].PutValue("Data D6");

                // ------------------------------------------------------------
                // 3. Apply a background color to the entire union range
                // ------------------------------------------------------------
                Style style = workbook.CreateStyle();
                style.Pattern = BackgroundType.Solid;
                style.ForegroundColor = Color.LightGreen; // desired background color

                // Apply the style to all formatting aspects of the union range
                unionRange.ApplyStyle(style, new StyleFlag { All = true });

                // ------------------------------------------------------------
                // 4. Save the workbook
                // ------------------------------------------------------------
                string outputPath = "UnionRangeNamedDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
