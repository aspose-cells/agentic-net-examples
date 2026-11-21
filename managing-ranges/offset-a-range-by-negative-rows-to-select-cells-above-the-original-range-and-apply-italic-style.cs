// Title: How to shift an Excel range upward with a negative row offset and apply italic font using Aspose.Cells for .NET (C#)
// Common Searches: c# aspocells offset a range upward by one row | aspocells apply italic font to a specific range | create a style flag for italic text in Aspose.Cells C# | move Excel range up using negative row offset in Aspose.Cells | apply formatting to an offset range in an Aspose.Cells workbook
// Tags: offset range upward Aspose.Cells C# | apply italic font style to range Aspose.Cells | style flag font italic Aspose.Cells | negative row offset Excel range C# | create offset range with CreateRange Aspose.Cells

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExample
{
    // The example creates a workbook, fills cells A2:B5, defines that range, computes a new range shifted one row upward (handling negative start rows), builds an italic font style, uses a StyleFlag to apply only the italic attribute, applies the style to the offset range, and saves the file as OffsetRangeItalic.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Fill sample data in the original range (A2:B5)
                sheet.Cells["A2"].PutValue("Item1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("Item2");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("Item3");
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["A5"].PutValue("Item4");
                sheet.Cells["B5"].PutValue(40);

                // Define the original range (A2:B5)
                AsposeRange originalRange = sheet.Cells.CreateRange("A2:B5");

                // Calculate the offset range (move up one row, same columns)
                int rowOffset = -1;
                int columnOffset = 0;
                int newFirstRow = originalRange.FirstRow + rowOffset;
                int newFirstColumn = originalRange.FirstColumn + columnOffset;

                // Ensure the new start row is not negative
                if (newFirstRow < 0) newFirstRow = 0;

                AsposeRange offsetRange = sheet.Cells.CreateRange(
                    newFirstRow,
                    newFirstColumn,
                    originalRange.RowCount,
                    originalRange.ColumnCount);

                // Create an italic style
                Style italicStyle = workbook.CreateStyle();
                italicStyle.Font.IsItalic = true;

                // Define which style attributes to apply (only font italic)
                StyleFlag styleFlag = new StyleFlag
                {
                    FontItalic = true
                };

                // Apply the italic style to the offset range
                offsetRange.ApplyStyle(italicStyle, styleFlag);

                // Save the workbook
                string outputPath = "OffsetRangeItalic.xlsx";
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
