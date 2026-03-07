using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Loading;

class DbfToExcelConverter
{
    static void Main()
    {
        // Path to the source DBF file (relative to the executable directory)
        string dbfPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "source.dbf");

        // Verify that the DBF file exists
        if (!File.Exists(dbfPath))
        {
            Console.WriteLine($"Error: DBF file not found at '{dbfPath}'.");
            return;
        }

        // Path for the resulting Excel file
        string excelPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "result.xlsx");

        // Create load options for DBF format
        DbfLoadOptions loadOptions = new DbfLoadOptions();

        // Load the DBF file into a workbook using the load options
        Workbook workbook = new Workbook(dbfPath, loadOptions);

        // Save the workbook as an Excel XLSX file
        workbook.Save(excelPath, SaveFormat.Xlsx);

        Console.WriteLine($"Conversion completed: '{dbfPath}' -> '{excelPath}'");
    }
}