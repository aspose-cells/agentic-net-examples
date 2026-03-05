using System;
using Aspose.Cells;

namespace AsposeCellsExportDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            ExportToXlsx.Run();
        }
    }

    public class ExportToXlsx
    {
        public static void Run()
        {
            Workbook workbook = new Workbook("input.xls");
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Exported to XLSX");
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook successfully exported to XLSX format.");
        }
    }
}