// Title: Shift a cell range up by one row and apply italic font using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a range, offsets it by -1 rows, and applies only the italic font style with Aspose.Cells. | Show how to compute a safe upward offset for an existing range and apply an italic StyleFlag in Aspose.Cells. | Generate a method that receives a worksheet and a source range, returns an offset range, and sets the font to italic.
// Common Searches: Aspose.Cells C# offset range upward by one row | apply italic style to a shifted range in Aspose.Cells | how to use StyleFlag to set only italic font with Aspose.Cells .NET | create an offset range without altering original cells Aspose.Cells C#
// Tags: offset range negative rows Aspose.Cells | apply italic StyleFlag Aspose.Cells | CreateRange method worksheet C# | cell formatting worksheet Aspose.Cells | save workbook as xlsx Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

namespace OffsetItalicExample
{
    // The example creates a workbook, defines range B3:C6, offsets it one row upward, applies an italic font style using a StyleFlag, and saves the file as OffsetItalicExample.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Define the original range (cells B3:C6)
                int firstRow = 2;        // zero‑based index (row 3)
                int firstColumn = 1;     // column B
                int totalRows = 4;
                int totalColumns = 2;
                AsposeRange originalRange = sheet.Cells.CreateRange(firstRow, firstColumn, totalRows, totalColumns);

                // Offset the range by -1 rows (select cells one row above)
                int offsetRows = -1;
                int offsetColumns = 0;   // no horizontal shift

                // Calculate new start position ensuring it stays within worksheet bounds
                int newFirstRow = Math.Max(firstRow + offsetRows, 0);
                int newFirstColumn = Math.Max(firstColumn + offsetColumns, 0);

                // Create the offset range using the new coordinates
                AsposeRange offsetRange = sheet.Cells.CreateRange(newFirstRow, newFirstColumn, totalRows, totalColumns);

                // Apply italic style to the offset range
                Style italicStyle = workbook.CreateStyle();
                italicStyle.Font.IsItalic = true;

                // Create a style flag to apply only the font italic property
                StyleFlag flag = new StyleFlag();
                flag.FontItalic = true;

                // Apply the style to the entire offset range
                offsetRange.ApplyStyle(italicStyle, flag);

                // Save the workbook
                string outputPath = "OffsetItalicExample.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
