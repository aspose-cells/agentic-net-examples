using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeCopyDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate source range (A1:E5) with sample data
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    cells[row, col].PutValue($"Data {row},{col}");
                }
            }

            // Define source and destination ranges
            // Source: rows 0-4, columns 0-4 (A1:E5)
            AsposeRange sourceRange = cells.CreateRange(0, 0, 5, 5);
            // Destination: rows 6-10, columns 0-4 (A7:E11)
            AsposeRange destinationRange = cells.CreateRange(6, 0, 5, 5);

            // Set paste options to copy everything (values, formulas, formats, etc.)
            PasteOptions pasteOptions = new PasteOptions
            {
                PasteType = PasteType.All
            };

            // Copy source range to destination range using the specified paste options
            destinationRange.Copy(sourceRange, pasteOptions);

            // Save the workbook in XLSX format
            workbook.Save("RangeCopyWithFormatting.xlsx");
        }
    }
}