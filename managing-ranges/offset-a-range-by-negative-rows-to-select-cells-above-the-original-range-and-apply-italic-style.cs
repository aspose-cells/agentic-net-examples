// Title: Apply Italic Font to an Upward‑Offset Range Using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, define a source range (A3:C4), shift it upward by two rows with GetOffset(-2,0) to target A1:C2, build an italic style, use a StyleFlag to affect only the font italic attribute, apply the style to the offset range, and save the file as OffsetRangeItalic.xlsx.
// Keywords: Aspose.Cells GetOffset | negative row offset | apply italic style | StyleFlag font italic | C# Aspose.Cells range formatting | offset range upward | Excel automation .NET
// Common Searches: Aspose.Cells offset range upward | How to apply only italic formatting to a range in C# | GetOffset method negative rows example | StyleFlag usage for font styles Aspose.Cells
// Developer Intent: Select cells above an existing range and change their font to italic without altering other formatting.
// Use Cases: Add a formatted header row above a data table. | Create a summary block above a report and emphasize text with italics. | Dynamically shift a range upward and apply italic styling for visual cues.
// AI Prompts: Generate C# code that offsets a range by -3 rows and applies both bold and italic formatting with Aspose.Cells. | Show how to offset a range upward and set a background color while preserving existing cell styles. | Explain the role of StyleFlag when applying only the italic attribute to a range in Aspose.Cells.

using System;
using Aspose.Cells;
using System.Drawing;

// Demonstrates how to create a workbook, define a source range (A3:C4), shift it upward by two rows with GetOffset(-2,0) to target A1:C2, build an italic style, use a StyleFlag to affect only the font italic attribute, apply the style to the offset range, and save the file as OffsetRangeItalic.xlsx.
class OffsetRangeApplyItalic
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate an original range (A3:C4) with sample data
            // A3:C4 corresponds to firstRow=2, firstColumn=0, totalRows=2, totalColumns=3
            Aspose.Cells.Range originalRange = cells.CreateRange(2, 0, 2, 3);
            for (int i = 0; i < originalRange.RowCount; i++)
            {
                for (int j = 0; j < originalRange.ColumnCount; j++)
                {
                    originalRange[i, j].PutValue($"R{2 + i}C{j + 1}");
                }
            }

            // Offset the range by -2 rows (select cells two rows above the original range)
            // This will point to range A1:C2
            Aspose.Cells.Range offsetRange = originalRange.GetOffset(-2, 0);

            // Create a style with italic font
            Style italicStyle = workbook.CreateStyle();
            italicStyle.Font.IsItalic = true;

            // Use StyleFlag to apply only the italic attribute
            StyleFlag flag = new StyleFlag();
            flag.FontItalic = true;

            // Apply the italic style to the offset range
            offsetRange.ApplyStyle(italicStyle, flag);

            // Save the workbook
            workbook.Save("OffsetRangeItalic.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
