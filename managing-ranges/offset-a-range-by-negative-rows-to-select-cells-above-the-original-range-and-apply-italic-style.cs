// Title: Shift a range up one row and apply italic formatting with Aspose.Cells for .NET (C#)
// Description: The sample builds a workbook, defines the range B2:D4, populates it with test data, moves the range one row higher using GetOffset(-1,0), creates an italic font style, applies it through a StyleFlag so only the italic attribute changes, and saves the result as OffsetItalicDemo.xlsx.
// Keywords: Aspose.Cells GetOffset | negative row offset | apply italic style C# | StyleFlag font italic | shift range upward Aspose.Cells | C# Excel formatting | Aspose.Cells range manipulation
// Common Searches: Aspose.Cells offset range negative rows | How to apply only italic font to a range in Aspose.Cells | GetOffset method example C# | StyleFlag usage Aspose.Cells | Shift range up one row Aspose.Cells .NET
// Developer Intent: Select cells located above an existing range and apply italic formatting without altering other style attributes.
// Use Cases: Create a header row above a data block and style it italic. | Add an annotation row above a table with italic text. | Generate a summary line above a report section while preserving original cell formatting.
// AI Prompts: Provide C# code to offset a range by -2 rows and apply both bold and italic using Aspose.Cells. | Show how to move a range left by two columns and set a background color while keeping cell values intact. | Explain the role of StyleFlag when applying multiple formatting options to an offset range in Aspose.Cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsOffsetItalicDemo
{
    // The sample builds a workbook, defines the range B2:D4, populates it with test data, moves the range one row higher using GetOffset(-1,0), creates an italic font style, applies it through a StyleFlag so only the italic attribute changes, and saves the result as OffsetItalicDemo.xlsx.
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

                // Populate sample data in the original range (B2:D4)
                AsposeRange originalRange = cells.CreateRange("B2", "D4");
                for (int i = 0; i < originalRange.RowCount; i++)
                {
                    for (int j = 0; j < originalRange.ColumnCount; j++)
                    {
                        originalRange[i, j].PutValue($"R{i + 2}C{j + 2}");
                    }
                }

                // Offset the original range by -1 row (one row above) and 0 columns
                AsposeRange offsetRange = originalRange.GetOffset(-1, 0);

                // Create a style with italic font
                Style italicStyle = workbook.CreateStyle();
                italicStyle.Font.IsItalic = true;

                // Use StyleFlag to apply only the italic attribute
                StyleFlag flag = new StyleFlag();
                flag.FontItalic = true;

                // Apply the italic style to the offset range
                offsetRange.ApplyStyle(italicStyle, flag);

                // Save the workbook
                string outputPath = "OffsetItalicDemo.xlsx";
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
