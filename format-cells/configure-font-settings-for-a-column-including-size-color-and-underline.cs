// Title: Configure column font size, color, underline and default style with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, target a column, define a Style with a 14‑pt blue font and single underline, use StyleFlag to apply only those font attributes, set the style as the column's default for future cells, and save the file as ColumnFontSettings.xlsx.
// Keywords: Aspose.Cells column font style C# | set column font size color underline | StyleFlag column formatting Aspose | default column style Aspose.Cells | C# Excel column formatting library
// Common Searches: Aspose.Cells change font size for entire column C# | apply blue underline to a column using Aspose.Cells | set default style for new cells in a column Aspose | StyleFlag usage for column formatting Aspose.Cells | C# Excel column font color and underline example
// Developer Intent: Apply a specific font size, color, and underline to all existing cells in a column and make that style the default for any new cells added to the column.
// Use Cases: Create a spreadsheet where the first column headings are uniformly styled (blue, 14 pt, underlined) without formatting each cell individually. | Ensure a data‑entry column automatically inherits a predefined font style for every new row. | Batch‑update an existing workbook to give a whole column a consistent appearance in a single operation.
// AI Prompts: Write C# code with Aspose.Cells to set a 12‑pt red font with double underline for column B and apply it as the default style for that column. | Explain the role of StyleFlag when applying column styles in Aspose.Cells and show an example that changes only font color and underline.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsColumnFontSettings
{
    // Demonstrates how to create a workbook, target a column, define a Style with a 14‑pt blue font and single underline, use StyleFlag to apply only those font attributes, set the style as the column's default for future cells, and save the file as ColumnFontSettings.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the first column (index 0)
            Column column = worksheet.Cells.Columns[0];

            // Create a style and configure the desired font properties
            Style fontStyle = workbook.CreateStyle();
            fontStyle.Font.Size = 14;                     // Font size
            fontStyle.Font.Color = Color.Blue;            // Font color
            fontStyle.Font.Underline = FontUnderlineType.Single; // Underline

            // Define which font attributes should be applied
            StyleFlag flag = new StyleFlag
            {
                FontSize = true,
                FontColor = true,
                FontUnderline = true
            };

            // Apply the style to existing cells in the column
            column.ApplyStyle(fontStyle, flag);

            // Set the style as the default for any new cells added to this column
            column.SetStyle(fontStyle);

            // Save the workbook
            workbook.Save("ColumnFontSettings.xlsx");
        }
    }
}
