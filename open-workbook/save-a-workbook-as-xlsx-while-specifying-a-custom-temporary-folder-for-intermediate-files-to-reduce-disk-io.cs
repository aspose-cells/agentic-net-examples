using System;
using Aspose.Cells;

namespace AsposeCellsCustomCacheDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // Specify a custom folder for temporary cache files
            string cacheFolder = @"C:\TempCache";

            // Create OOXML save options for XLSX format
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.CachedFileFolder = cacheFolder; // Reduce disk I/O by using the custom folder

            // Save the workbook as XLSX using the save options
            workbook.Save("output.xlsx", saveOptions);

            Console.WriteLine("Workbook saved as XLSX with custom cache folder.");
        }
    }
}