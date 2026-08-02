// Title: Apply a Two‑Level Indent Style to a Nested List Range with Aspose.Cells for .NET
// Description: This example creates a workbook, fills cells A1‑A5 with hierarchical items, defines a style with IndentLevel = 2 and left alignment, enables the indent flag, builds a range covering A1:A5, applies the style to the whole range, and saves the file as NestedListIndent.xlsx.
// Keywords: Aspose.Cells | .NET | C# indent style | cell indentation | IndentLevel | HorizontalAlignment left | StyleFlag Indent | Range.ApplyStyle | nested list Excel | Excel formatting automation
// Common Searches: Aspose.Cells set indent level for a range | apply indentation to multiple cells C# Aspose.Cells | create hierarchical list with indent in Excel using Aspose | IndentLevel not working Aspose.Cells | how to use StyleFlag with indent in Aspose.Cells
// Developer Intent: Create a style that indents cells by two levels and apply it to a range representing a nested list in an Excel worksheet.
// Use Cases: Display parent‑child relationships in a single column with visual indentation. | Generate a table of contents where each level is offset uniformly. | Produce reports where sub‑sections need consistent left indentation across rows.
// AI Prompts: Generate C# code with Aspose.Cells that sets a three‑level indent style and applies it only to rows 2‑6. | Explain why HorizontalAlignment must be set to Left for IndentLevel to take effect in Aspose.Cells. | Show how to apply different indent levels to several separate ranges in the same worksheet using Aspose.Cells. | Provide a step‑by‑step guide to create and reuse an indentation style across multiple workbooks with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsIndentExample
{
    // This example creates a workbook, fills cells A1‑A5 with hierarchical items, defines a style with IndentLevel = 2 and left alignment, enables the indent flag, builds a range covering A1:A5, applies the style to the whole range, and saves the file as NestedListIndent.xlsx.
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

                // Populate sample data representing a nested list
                cells["A1"].PutValue("Item 1");
                cells["A2"].PutValue("Item 1.1");
                cells["A3"].PutValue("Item 1.1.1");
                cells["A4"].PutValue("Item 2");
                cells["A5"].PutValue("Item 2.1");

                // Create a style with an indent level of 2
                Style indentStyle = workbook.CreateStyle();
                indentStyle.IndentLevel = 2; // two indentation levels
                indentStyle.HorizontalAlignment = TextAlignmentType.Left; // required for indent to work

                // Indicate that the indent setting should be applied
                StyleFlag flag = new StyleFlag { Indent = true };

                // Define the range that represents the nested list (A1:A5)
                AsposeRange nestedListRange = cells.CreateRange("A1", "A5");

                // Apply the indent style to the entire range
                nestedListRange.ApplyStyle(indentStyle, flag);

                // Determine output file path
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "NestedListIndent.xlsx");

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
