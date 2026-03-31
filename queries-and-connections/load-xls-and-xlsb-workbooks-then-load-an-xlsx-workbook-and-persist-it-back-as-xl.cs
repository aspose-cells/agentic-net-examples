using System;
using Aspose.Cells;

namespace AsposeCellsLoadAndSaveDemo
{
    class Program
    {
        static void Main()
        {
            // Load an Excel 97-2003 workbook (XLS)
            Workbook xlsWorkbook = new Workbook("input.xls");
            // At this point the workbook is loaded and can be used for further processing

            // Load an Excel Binary Workbook (XLSB)
            Workbook xlsbWorkbook = new Workbook("input.xlsb");
            // The workbook is now in memory

            // Load an Excel Open XML workbook (XLSX)
            Workbook xlsxWorkbook = new Workbook("input.xlsx");

            // Persist the loaded XLSX workbook back to disk as XLSX
            // The Save method with a file name and SaveFormat ensures the correct format is used
            xlsxWorkbook.Save("output.xlsx", SaveFormat.Xlsx);

            // Optional: release resources
            xlsWorkbook.Dispose();
            xlsbWorkbook.Dispose();
            xlsxWorkbook.Dispose();

            Console.WriteLine("XLSX workbook loaded and saved successfully.");
        }
    }
}