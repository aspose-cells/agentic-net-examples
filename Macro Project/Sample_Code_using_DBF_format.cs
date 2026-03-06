using System;
using Aspose.Cells;
using Aspose.Cells.Saving;
using Aspose.Cells.Loading;

namespace AsposeCellsDbfSample
{
    class Program
    {
        static void Main()
        {
            // -------------------- Create and populate a workbook --------------------
            Workbook workbook = new Workbook();                     // create a new workbook
            Worksheet sheet = workbook.Worksheets[0];              // get the first worksheet

            // Add sample data
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("John");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Alice");

            // -------------------- Save as DBF --------------------
            DbfSaveOptions saveOptions = new DbfSaveOptions();    // create DBF save options
            saveOptions.ExportAsString = true;                     // export all values as strings (optional)
            string dbfPath = "sample.dbf";
            workbook.Save(dbfPath, saveOptions);                  // save workbook in DBF format

            Console.WriteLine($"Workbook saved to DBF file: {dbfPath}");

            // -------------------- Load the DBF file --------------------
            DbfLoadOptions loadOptions = new DbfLoadOptions();    // create DBF load options
            Workbook loadedWorkbook = new Workbook(dbfPath, loadOptions); // load DBF file

            // Access loaded data
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            Console.WriteLine("Data loaded from DBF:");
            Console.WriteLine($"Cell A2: {loadedSheet.Cells["A2"].StringValue}");
            Console.WriteLine($"Cell B2: {loadedSheet.Cells["B2"].StringValue}");

            // Modify a cell to demonstrate further processing
            loadedSheet.Cells["B2"].PutValue("John Updated");

            // -------------------- Save the modified workbook as XLSX --------------------
            string xlsxPath = "modified.xlsx";
            loadedWorkbook.Save(xlsxPath, SaveFormat.Xlsx);
            Console.WriteLine($"Modified workbook saved as XLSX: {xlsxPath}");
        }
    }
}