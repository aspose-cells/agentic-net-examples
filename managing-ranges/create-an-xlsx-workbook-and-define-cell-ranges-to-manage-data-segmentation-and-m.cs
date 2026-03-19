using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        for (int i = 0; i < 5; i++)
        {
            cells[i, 0].PutValue(i + 1);
            cells[i, 1].PutValue((i + 1) * 10);
        }

        AsposeRange dataRange = cells.CreateRange("A1", "B5");
        dataRange.Name = "DataRange";

        cells.InsertRow(2);
        cells.InsertColumn(1);

        cells["C1"].Formula = "=SUM(DataRange)";

        workbook.CalculateFormula();

        Console.WriteLine("Sum of DataRange: " + cells["C1"].IntValue);

        workbook.Save("SegmentedData.xlsx");
    }
}