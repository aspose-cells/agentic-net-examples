using System;
using System.Data;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a DataTable with HTML content
        DataTable table = new DataTable();
        table.Columns.Add("ID", typeof(int));
        table.Columns.Add("Content", typeof(string));
        table.Rows.Add(1, "<a href='https://example.com'>Example Link</a>");
        table.Rows.Add(2, "<b>Bold Text</b> and <i>Italic Text</i>");

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure import options: do NOT preserve HTML (strip tags)
        ImportTableOptions importOptions = new ImportTableOptions
        {
            IsHtmlString = false,      // Treat values as plain text, remove HTML tags
            IsFieldNameShown = true    // Import column headers as the first row
        };

        // Import the DataTable into the worksheet starting at cell A1 (row 0, column 0)
        worksheet.Cells.ImportData(table, 0, 0, importOptions);

        // Verify that HTML tags have been removed from the imported cells
        // Column B (index 1) contains the "Content" values
        for (int row = 0; row < table.Rows.Count + 1; row++) // +1 for header row
        {
            string cellText = worksheet.Cells[row, 1].StringValue; // B column
            bool containsHtml = cellText.Contains("<") || cellText.Contains(">");
            Console.WriteLine($"Cell B{row + 1}: \"{cellText}\" {(containsHtml ? "contains HTML tags!" : "clean")}");
        }

        // Save the workbook to verify the result
        workbook.Save("ImportHtmlRemoved.xlsx");
    }
}