using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsTimestampImport
{
    class Program
    {
        static void Main()
        {
            // Sample timestamps to import
            DateTime[] timestamps = new DateTime[]
            {
                new DateTime(2023, 5, 1, 8, 15, 30),
                new DateTime(2023, 5, 1, 12, 45, 0),
                new DateTime(2023, 5, 1, 17, 5, 9)
            };

            // Prepare a DataTable with a single DateTime column
            DataTable table = new DataTable();
            table.Columns.Add("TimeStamp", typeof(DateTime));
            foreach (DateTime ts in timestamps)
            {
                table.Rows.Add(ts);
            }

            // Create a workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Set import options: apply time format "HH:mm:ss" to the first (and only) column
            ImportTableOptions importOptions = new ImportTableOptions
            {
                // NumberFormats aligns with column index; null for columns we don't format
                NumberFormats = new string[] { "HH:mm:ss" }
            };

            // Import the DataTable starting at cell A1 (row 0, column 0)
            worksheet.Cells.ImportData(table, 0, 0, importOptions);

            // Save the workbook
            workbook.Save("TimestampsFormatted.xlsx");
        }
    }
}