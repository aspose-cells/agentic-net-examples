using System;
using Aspose.Cells;
using Aspose.Cells.Saving;
using Aspose.Cells.Loading;

namespace AsposeCellsDbfConsoleDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("John");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Alice");

            // Save the workbook as a DBF file using DbfSaveOptions
            DbfSaveOptions saveOptions = new DbfSaveOptions();
            saveOptions.ExportAsString = true; // export all values as strings
            string dbfPath = "SampleData.dbf";
            workbook.Save(dbfPath, saveOptions);
            Console.WriteLine($"Workbook saved to DBF file: {dbfPath}");

            // Load the DBF file back using DbfLoadOptions
            DbfLoadOptions loadOptions = new DbfLoadOptions();
            Workbook loadedWorkbook = new Workbook(dbfPath, loadOptions);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

            // Output the loaded data to the console
            Console.WriteLine("Data loaded from DBF file:");
            int maxRow = loadedSheet.Cells.MaxDataRow;
            int maxCol = loadedSheet.Cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Console.Write(loadedSheet.Cells[row, col].StringValue);
                    if (col < maxCol) Console.Write("\t");
                }
                Console.WriteLine();
            }
        }
    }
}