using System;
using Aspose.Cells;
using Aspose.Cells.Saving;
using Aspose.Cells.Loading;

class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("ID");
        worksheet.Cells["B1"].PutValue("Name");
        worksheet.Cells["A2"].PutValue(1);
        worksheet.Cells["B2"].PutValue("John");
        worksheet.Cells["A3"].PutValue(2);
        worksheet.Cells["B3"].PutValue("Alice");

        // Create DBF save options and export all values as strings
        DbfSaveOptions saveOptions = new DbfSaveOptions();
        saveOptions.ExportAsString = true;

        // Save the workbook as a DBF file
        string dbfFile = "sample.dbf";
        workbook.Save(dbfFile, saveOptions);

        // Load the DBF file using default load options
        DbfLoadOptions loadOptions = new DbfLoadOptions();
        Workbook dbfWorkbook = new Workbook(dbfFile, loadOptions);
        Worksheet dbfWorksheet = dbfWorkbook.Worksheets[0];

        // Determine the used range of the sheet
        int maxRow = dbfWorksheet.Cells.MaxDataRow;
        int maxCol = dbfWorksheet.Cells.MaxDataColumn;

        // Render field definitions and records to the console
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Console.Write(dbfWorksheet.Cells[row, col].StringValue);
                if (col < maxCol) Console.Write("\t");
            }
            Console.WriteLine();
        }
    }
}