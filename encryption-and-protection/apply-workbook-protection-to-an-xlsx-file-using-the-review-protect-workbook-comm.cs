using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ProtectWorkbookDemo
    {
        public static void Run()
        {
            Workbook workbook = new Workbook("input.xlsx");
            workbook.Protect(ProtectionType.Structure, "myPassword123");
            workbook.Save("protected_workbook.xlsx", SaveFormat.Xlsx);
            workbook.Dispose();
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ProtectWorkbookDemo.Run();
        }
    }
}