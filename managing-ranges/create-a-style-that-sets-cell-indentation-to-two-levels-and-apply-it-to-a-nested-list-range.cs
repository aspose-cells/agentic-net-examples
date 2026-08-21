// Title: Aspose.Cells .NET: Apply a Two‑Level Indent Style to a Nested List Range
// Description: Creates a workbook, defines a style with IndentLevel = 2 and left alignment, uses a StyleFlag to apply only the indent, selects range A2:A5 (a nested list), applies the style, and saves the file as XLSX.
// Keywords: Aspose.Cells | C# indent level | StyleFlag | ApplyStyle | Excel indentation | nested list formatting | range styling | CreateStyle | HorizontalAlignment | SaveFormat.Xlsx
// Common Searches: Aspose.Cells set indent level for a range | How to apply indentation to cells in Aspose.Cells C# | Create style with indent flag Aspose.Cells | Indent hierarchical data in generated Excel file | Apply style to range A2:A5 Aspose.Cells
// Developer Intent: Create a two‑level indent style and apply it exclusively to a specific cell range that represents a nested list.
// Use Cases: Visually represent parent‑child relationships in exported reports. | Format outline levels for task or item lists in automated Excel generation. | Maintain existing cell formatting while adding consistent left‑aligned indentation to sub‑items.
// AI Prompts: Generate C# code with Aspose.Cells that sets IndentLevel = 3, uses right alignment, and applies the style to range B2:B10 while preserving other formatting. | Show how to combine the two‑level indent style with bold font for first‑level items in a nested list. | Explain how to reuse the created indent style across multiple worksheets in the same workbook.

using System;
using Aspose.Cells;

// Creates a workbook, defines a style with IndentLevel = 2 and left alignment, uses a StyleFlag to apply only the indent, selects range A2:A5 (a nested list), applies the style, and saves the file as XLSX.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data representing a nested list
            cells["A1"].PutValue("Item");
            cells["A2"].PutValue("SubItem 1");
            cells["A3"].PutValue("SubItem 2");
            cells["A4"].PutValue("SubSubItem 1");
            cells["A5"].PutValue("SubSubItem 2");

            // Create a style and set the indent level to two
            Style indentStyle = workbook.CreateStyle();
            indentStyle.IndentLevel = 2;
            // Indent works only with left or right alignment
            indentStyle.HorizontalAlignment = TextAlignmentType.Left;

            // Create a style flag indicating that only the indent should be applied
            StyleFlag flag = new StyleFlag();
            flag.Indent = true;

            // Define the range that contains the nested list (A2:A5)
            Aspose.Cells.Range nestedRange = cells.CreateRange("A2", "A5");

            // Apply the indent style to the range
            nestedRange.ApplyStyle(indentStyle, flag);

            // Save the workbook
            workbook.Save("NestedListIndent.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
