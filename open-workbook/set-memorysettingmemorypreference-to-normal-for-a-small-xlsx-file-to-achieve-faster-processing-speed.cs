using System;
using Aspose.Cells;

namespace AsposeCellsMemorySettingDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (small XLSX file)
            Workbook workbook = new Workbook();

            // Set the memory usage option to Normal for faster processing on small files
            workbook.Settings.MemorySetting = MemorySetting.Normal;

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Memory Setting Demo");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Save the workbook to an XLSX file
            workbook.Save("MemorySettingNormalDemo.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook saved with MemorySetting.Normal.");
        }
    }
}