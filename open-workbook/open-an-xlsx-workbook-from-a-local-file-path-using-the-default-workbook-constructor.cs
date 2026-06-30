using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Author note: example to open an XLSX workbook from a local file path
        string filePath = "sample.xlsx";
        Workbook workbook = new Workbook(filePath);
        Console.WriteLine("Number of worksheets: " + workbook.Worksheets.Count);
    }
}