using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Sample data (ensures no runtime failure)
        worksheet.Cells["A1"].PutValue("Name");
        worksheet.Cells["B1"].PutValue("Age");
        worksheet.Cells["A2"].PutValue("Ali");
        worksheet.Cells["B2"].PutValue(25);
        worksheet.Cells["A3"].PutValue("Sara");
        worksheet.Cells["B3"].PutValue(30);

        int lastRow = worksheet.Cells.MaxDataRow;
        int lastColumn = worksheet.Cells.MaxDataColumn;

        if (lastRow < 0 || lastColumn < 0)
        {
            Console.WriteLine("No data found.");
            return;
        }

        int tableIndex = worksheet.ListObjects.Add(0, 0, lastRow, lastColumn, true);
        ListObject listObject = worksheet.ListObjects[tableIndex];

        listObject.DisplayName = "DataTable";

        workbook.Save("output.xlsx");

        Console.WriteLine("Table created successfully.");
        Console.ReadLine();
    }
}