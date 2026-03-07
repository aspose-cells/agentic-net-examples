using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTimelineConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel template (XLTX) that contains a Timeline control
            string sourcePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TimelineTemplate.xltx");

            // Desired output path for the converted workbook (XLSX)
            string destinationPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TimelineWorkbook.xlsx");

            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the XLTX template and save it as XLSX
            var workbook = new Workbook(sourcePath);
            workbook.Save(destinationPath, SaveFormat.Xlsx);

            Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{destinationPath}'");
        }
    }
}