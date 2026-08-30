// Title: Import a DateTime array into an Excel worksheet and format the cells as HH:mm:ss using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a Workbook, imports a DateTime[] into the first worksheet, and applies the "hh:mm:ss" number format to the imported column with Aspose.Cells. | Show how to use ImportTableOptions to set a custom time format while importing a DataTable of timestamps into an Excel file using Aspose.Cells. | Generate a complete C# console program that saves the formatted timestamps to a .xlsx file named TimestampsFormatted.xlsx with Aspose.Cells.
// Common Searches: asp.net import datetime array into excel and display time as hh:mm:ss using aspose.cells | c# aspose.cells import datatable of timestamps with custom time format | how to set number format hh:mm:ss when importing data with aspose.cells | save timestamps to xlsx with 24‑hour time format using aspose.cells c#
// Tags: ImportTableOptions time number format | C# Aspose.Cells import DateTime array | format cells hh:mm:ss Excel | save workbook as xlsx Aspose.Cells | DataTable timestamp import Aspose.Cells

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsTimestampImport
{
    // Creates a new Workbook, builds a DataTable with DateTime values, imports it into the first worksheet using ImportTableOptions that specify the "hh:mm:ss" number format, and saves the workbook as TimestampsFormatted.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Prepare a DataTable that holds timestamp values
            DataTable timeTable = new DataTable();
            timeTable.Columns.Add("Timestamp", typeof(DateTime));

            // Example timestamps – you can replace these with your own array values
            DateTime[] timestamps = new DateTime[]
            {
                new DateTime(2023, 1, 1, 8, 15, 30),
                new DateTime(2023, 1, 1, 12, 45, 5),
                new DateTime(2023, 1, 1, 23, 59, 59)
            };

            foreach (DateTime ts in timestamps)
            {
                timeTable.Rows.Add(ts);
            }

            // Define import options and set the number format for the first (and only) column
            ImportTableOptions importOptions = new ImportTableOptions
            {
                // "hh:mm:ss" displays time in 24‑hour format with leading zeros
                NumberFormats = new string[] { "hh:mm:ss" },
                // Optional: show column header if desired
                IsFieldNameShown = false
            };

            // Import the DataTable into the worksheet starting at cell A1 (row 0, column 0)
            worksheet.Cells.ImportData(timeTable, 0, 0, importOptions);

            // Save the workbook
            workbook.Save("TimestampsFormatted.xlsx");
        }
    }
}
