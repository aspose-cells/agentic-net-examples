using System;
using System.Data;
using Aspose.Cells;

namespace TimestampImportExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define an array of timestamps (DateTime values)
            DateTime[] timestamps = new DateTime[]
            {
                new DateTime(2023, 1, 1, 8, 15, 30),
                new DateTime(2023, 1, 1, 12, 45, 0),
                new DateTime(2023, 1, 1, 16, 30, 15),
                new DateTime(2023, 1, 1, 20, 5, 5)
            };

            // Prepare a DataTable with a single DateTime column to hold the timestamps
            DataTable table = new DataTable();
            table.Columns.Add("Timestamp", typeof(DateTime));
            foreach (DateTime ts in timestamps)
            {
                table.Rows.Add(ts);
            }

            // Set import options:
            // - Do not import column header (IsFieldNameShown = false)
            // - Apply a custom number format "hh:mm:ss" to the first column
            ImportTableOptions importOptions = new ImportTableOptions
            {
                IsFieldNameShown = false,
                NumberFormats = new string[] { "hh:mm:ss" }   // format for the first column
            };

            // Import the DataTable starting at cell A1 (row 0, column 0)
            worksheet.Cells.ImportData(table, 0, 0, importOptions);

            // Save the workbook
            workbook.Save("Timestamps.xlsx");
        }
    }
}