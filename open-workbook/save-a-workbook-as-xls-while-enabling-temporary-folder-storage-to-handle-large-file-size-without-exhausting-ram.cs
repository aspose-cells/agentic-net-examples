using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLargeFileExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Value");
            for (int i = 2; i <= 100000; i++)
            {
                sheet.Cells[$"A{i}"].PutValue(i - 1);
                sheet.Cells[$"B{i}"].PutValue($"Data {i - 1}");
            }

            // Define a temporary folder for caching large data during save
            string cacheFolder = Path.Combine(Path.GetTempPath(), "AsposeCache");
            Directory.CreateDirectory(cacheFolder);

            // Configure XlsSaveOptions to use the temporary cache folder
            XlsSaveOptions saveOptions = new XlsSaveOptions();
            saveOptions.CachedFileFolder = cacheFolder;

            // Save the workbook as an Excel 97-2003 XLS file using the configured options
            workbook.Save("LargeOutput.xls", saveOptions);

            Console.WriteLine("Workbook saved as XLS with temporary cache folder enabled.");
        }
    }
}