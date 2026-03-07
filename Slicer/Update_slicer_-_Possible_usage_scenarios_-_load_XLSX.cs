using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        ws.Cells["A1"].PutValue("Hello Aspose.Cells!");

        // Save the workbook
        wb.Save("output.xlsx");
        Console.WriteLine("Workbook created.");
    }
}