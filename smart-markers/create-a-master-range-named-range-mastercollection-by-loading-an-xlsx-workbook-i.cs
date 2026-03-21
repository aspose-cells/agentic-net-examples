using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        string filePath = "input.xlsx";
        Workbook workbook = new Workbook(filePath);
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        var masterRange = cells.CreateRange("A1:Z100");
        masterRange.Name = "MasterCollection";
    }
}