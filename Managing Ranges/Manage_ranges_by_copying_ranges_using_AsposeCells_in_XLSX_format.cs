using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        AsposeRange sourceRange = cells.CreateRange(0, 0, 3, 3);
        AsposeRange destFullCopy = cells.CreateRange(5, 0, 3, 3);
        AsposeRange destValuesOnly = cells.CreateRange(9, 0, 3, 3);
        AsposeRange destFormatsOnly = cells.CreateRange(13, 0, 3, 3);

        destFullCopy.Copy(sourceRange);
        destValuesOnly.CopyValue(sourceRange);

        PasteOptions pasteOptions = new PasteOptions
        {
            PasteType = PasteType.Formats
        };
        destFormatsOnly.Copy(sourceRange, pasteOptions);

        workbook.Save("RangeCopyDemo.xlsx");
    }
}