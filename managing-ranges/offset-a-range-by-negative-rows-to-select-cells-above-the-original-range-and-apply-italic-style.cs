using System;
using Aspose.Cells;
using System.Drawing;
using System.IO;

namespace AsposeCellsOffsetItalicDemo
{
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
                Aspose.Cells.Range originalRange = cells.CreateRange("B2", "D4");
                for (int i = 0; i < originalRange.RowCount; i++)
                {
                    for (int j = 0; j < originalRange.ColumnCount; j++)
                    {
                        originalRange[i, j].PutValue($"R{i + 2}C{j + 2}");
                    }
                }

                // Offset the range by -2 rows (two rows above the original range)
                Aspose.Cells.Range offsetRange = originalRange.GetOffset(-2, 0);

                // Create a style with italic font
                Style italicStyle = workbook.CreateStyle();
                italicStyle.Font.IsItalic = true;

                // Apply only the italic attribute using StyleFlag
                StyleFlag flag = new StyleFlag();
                flag.FontItalic = true;

                // Apply the italic style to the offset range
                offsetRange.ApplyStyle(italicStyle, flag);

                // Save the workbook
                string outputPath = "OffsetItalicDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}