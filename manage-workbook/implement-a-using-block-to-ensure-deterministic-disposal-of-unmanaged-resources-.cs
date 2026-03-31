using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        using (Workbook workbook = new Workbook())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello Aspose!");
            workbook.Save("output.xlsx");
        }
    }
}