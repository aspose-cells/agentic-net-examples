using System;
using Aspose.Cells;

namespace TimelineConversionExample
{
    class Program
    {
        static void Main()
        {
            string sourcePath = "TimelineData.xls";
            string destinationPath = "TimelineData.xlsx";

            // Ensure the source file exists; create a simple workbook if it does not.
            if (!System.IO.File.Exists(sourcePath))
            {
                var wb = new Workbook();
                var ws = wb.Worksheets[0];
                ws.Name = "Sheet1";
                ws.Cells["A1"].PutValue("Sample Data");
                wb.Save(sourcePath, SaveFormat.Excel97To2003);
            }

            // Load the source workbook and save it as XLSX.
            var workbook = new Workbook(sourcePath);
            workbook.Save(destinationPath, SaveFormat.Xlsx);

            Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{destinationPath}'");
        }
    }
}