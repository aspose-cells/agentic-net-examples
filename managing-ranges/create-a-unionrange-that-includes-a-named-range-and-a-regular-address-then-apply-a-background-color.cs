// Title: Aspose.Cells .NET – Build a UnionRange from a named range and a cell address and set a solid background color
// Description: Demonstrates how to create a workbook, define a named range (A1:A5), populate it, combine it with another range (C1:C5) using Workbook.Worksheets.CreateUnionRange, and apply a solid light‑green fill to the entire UnionRange with a StyleFlag. The result is saved as UnionRangeNamedDemo.xlsx.
// Keywords: Aspose.Cells UnionRange C# | named range union Aspose.Cells | apply background color UnionRange | CreateUnionRange example | .NET Excel styling | solid fill style Aspose.Cells
// Common Searches: Aspose.Cells create UnionRange from named range and address | how to apply background color to UnionRange in C# | combine named range with another range Aspose.Cells | style entire UnionRange Aspose.Cells .NET
// Developer Intent: Combine a predefined named range with an additional cell range into a UnionRange and apply uniform formatting.
// Use Cases: Apply the same highlight to non‑contiguous sections (e.g., a named range and a separate column) in a generated report. | Quickly format multiple disjoint ranges with a single style when exporting data to Excel. | Maintain consistent visual cues across named and ad‑hoc ranges in automated spreadsheet creation.
// AI Prompts: Generate C# code that builds a UnionRange from a named range and a regular address and fills it with a solid light‑green background using Aspose.Cells. | Explain step‑by‑step how Workbook.Worksheets.CreateUnionRange can merge a named range and a cell range, then style the whole area with a StyleFlag. | Provide instructions to define a named range, create a UnionRange that includes another range, and apply a solid fill to all cells in the UnionRange.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsUnionRangeDemo
{
    // Demonstrates how to create a workbook, define a named range (A1:A5), populate it, combine it with another range (C1:C5) using Workbook.Worksheets.CreateUnionRange, and apply a solid light‑green fill to the entire UnionRange with a StyleFlag. The result is saved as UnionRangeNamedDemo.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Define a regular range (A1:A5) and give it a name
                Aspose.Cells.Range namedRange = worksheet.Cells.CreateRange("A1:A5");
                namedRange.Name = "MyRange";

                // Populate the named range with sample data
                for (int i = 0; i < namedRange.RowCount; i++)
                {
                    namedRange[i, 0].PutValue($"Item {i + 1}");
                }

                // Create a union range that includes the named range and another address (C1:C5)
                UnionRange unionRange = workbook.Worksheets.CreateUnionRange("MyRange,C1:C5", 0);

                // Apply a solid background color to the entire union range
                Style style = workbook.CreateStyle();
                style.Pattern = BackgroundType.Solid;
                style.ForegroundColor = Color.LightGreen;

                // Apply the style to all formatting aspects of the union range
                unionRange.ApplyStyle(style, new StyleFlag { All = true });

                // Save the workbook
                workbook.Save("UnionRangeNamedDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
