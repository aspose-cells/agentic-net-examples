using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsHtmlImportDemo
{
    class Program
    {
        static void Main()
        {
            // Create a DataTable containing HTML formatted text
            DataTable table = new DataTable();
            table.Columns.Add("ID");
            table.Columns.Add("Content");
            table.Columns.Add("Time");

            // Sample row with an HTML anchor tag
            table.Rows.Add("1", "<a href='https://www.example.com'>Example Link</a>", "2:30 PM");

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure import options:
            // - Show field names (column headers)
            // - Preserve HTML formatting by setting IsHtmlString to true
            // - Apply a time format to the third column
            ImportTableOptions importOptions = new ImportTableOptions
            {
                IsFieldNameShown = true,
                IsHtmlString = true,
                NumberFormats = new string[] { null, null, "h:mm AM/PM" }
            };

            // Import the DataTable into the worksheet starting at cell A1 (row 0, column 0)
            worksheet.Cells.ImportData(table, 0, 0, importOptions);

            // Save the workbook; the HTML content will be retained in the cell as rich text
            workbook.Save("HtmlStringImportDemo.xlsx");

            Console.WriteLine("Import completed. Workbook saved as HtmlStringImportDemo.xlsx");
        }
    }
}