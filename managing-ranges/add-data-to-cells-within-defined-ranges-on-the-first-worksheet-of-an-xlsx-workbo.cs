using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        Aspose.Cells.Range range = cells.CreateRange("A1", "C3");

        for (int i = 0; i < range.RowCount; i++)
        {
            for (int j = 0; j < range.ColumnCount; j++)
            {
                range[i, j].PutValue($"R{i + 1}C{j + 1}");
            }
        }

        workbook.Save("AddedData.xlsx", SaveFormat.Xlsx);
    }
}