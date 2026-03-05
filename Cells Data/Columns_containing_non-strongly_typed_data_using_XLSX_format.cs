using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsMixedColumnDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with mixed data types in the same column
            // Column A will contain int, double, and string values
            sheet.Cells["A1"].PutValue("ID");          // Header
            sheet.Cells["A2"].PutValue(1);            // Integer
            sheet.Cells["A3"].PutValue(2.5);          // Double
            sheet.Cells["A4"].PutValue("3");          // String that looks like a number
            sheet.Cells["A5"].PutValue(DateTime.Now); // DateTime

            // Column B will contain dates and strings
            sheet.Cells["B1"].PutValue("Date");
            sheet.Cells["B2"].PutValue(DateTime.Today);
            sheet.Cells["B3"].PutValue("2023-01-15"); // String date
            sheet.Cells["B4"].PutValue(DateTime.Now.AddDays(1));
            sheet.Cells["B5"].PutValue("InvalidDate"); // Non‑date string

            // Configure export options to check for mixed value types
            ExportTableOptions exportOptions = new ExportTableOptions
            {
                ExportColumnName = true,      // Use first row as column names
                CheckMixedValueType = true    // Examine all rows to determine column type
            };

            // Export the range (including header) to a DataTable
            DataTable dt = sheet.Cells.ExportDataTable(0, 0, 5, 2, exportOptions);

            // Display the inferred data types of each column
            Console.WriteLine("Column Types after Export:");
            foreach (DataColumn col in dt.Columns)
            {
                Console.WriteLine($"{col.ColumnName}: {col.DataType}");
            }

            // Save the workbook in XLSX format
            workbook.Save("MixedColumnsDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}