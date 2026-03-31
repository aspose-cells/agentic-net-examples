using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    public class LoadWorkbookDemo
    {
        public static void Run()
        {
            string filePath = "input.xlsx";
            Workbook workbook = new Workbook(filePath);
            Worksheet worksheet = workbook.Worksheets[0];
            Console.WriteLine("Cell A1 value: " + worksheet.Cells["A1"].StringValue);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            LoadWorkbookDemo.Run();
        }
    }
}