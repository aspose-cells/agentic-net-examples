using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class GetMaxDisplayRange
{
    static void Main()
    {
        // Create a new workbook (XLSX format by default)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data to create a display range
        worksheet.Cells["A1"].PutValue("Header1");
        worksheet.Cells["B1"].PutValue("Header2");
        worksheet.Cells["A2"].PutValue(100);
        worksheet.Cells["B2"].PutValue(200);
        worksheet.Cells["A3"].PutValue(300);
        worksheet.Cells["B3"].PutValue(400);

        // Get the maximum display range (includes data, merged cells, and shapes)
        AsposeRange maxDisplayRange = worksheet.Cells.MaxDisplayRange;

        // Output range information
        Console.WriteLine("Max Display Range:");
        Console.WriteLine($"Start Row: {maxDisplayRange.FirstRow}");
        Console.WriteLine($"Start Column: {maxDisplayRange.FirstColumn}");
        Console.WriteLine($"Total Rows: {maxDisplayRange.RowCount}");
        Console.WriteLine($"Total Columns: {maxDisplayRange.ColumnCount}");

        // Save the workbook as XLSX
        workbook.Save("MaxDisplayRangeDemo.xlsx");
    }
}