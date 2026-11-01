// Title: How to offset a cell range upward by two rows and apply italic font style with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# offset range upward by negative rows example | apply italic font to a specific range using Aspose.Cells .NET | create a shifted range above original cells in Aspose.Cells | using StyleFlag to set only italic attribute in Aspose.Cells | prevent out‑of‑bounds when offsetting ranges in Aspose.Cells
// Tags: negative row offset range Aspose.Cells | italic font style flag Aspose.Cells | range bounds check Aspose.Cells | create range from coordinates Aspose.Cells | apply style to offset range Aspose.Cells

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example creates a workbook, defines range B3:D5, offsets it two rows upward while staying inside worksheet limits, builds an italic font style, applies only the italic attribute to the offset range using a StyleFlag, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the original range (example: B3:D5)
            AsposeRange originalRange = sheet.Cells.CreateRange("B3:D5");

            // Calculate the offset (‑2 rows, 0 columns)
            int offsetRows = -2;
            int offsetColumns = 0;

            // Determine the new range coordinates
            int newFirstRow = originalRange.FirstRow + offsetRows;
            int newFirstColumn = originalRange.FirstColumn + offsetColumns;
            int rowCount = originalRange.RowCount;
            int columnCount = originalRange.ColumnCount;

            // Ensure the new range stays within worksheet bounds
            if (newFirstRow < 0) newFirstRow = 0;
            if (newFirstColumn < 0) newFirstColumn = 0;

            // Create the offset range
            AsposeRange offsetRange = sheet.Cells.CreateRange(newFirstRow, newFirstColumn, rowCount, columnCount);

            // Create a style with italic font
            Style italicStyle = workbook.CreateStyle();
            italicStyle.Font.IsItalic = true;

            // Apply only the italic attribute to the offset range
            StyleFlag flag = new StyleFlag { FontItalic = true };
            offsetRange.ApplyStyle(italicStyle, flag);

            // Save the workbook
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
