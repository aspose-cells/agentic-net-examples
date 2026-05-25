using System;
using System.Data;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet's cells
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Prepare a DataTable that contains HTML tags in its data
        DataTable table = new DataTable();
        table.Columns.Add("ID", typeof(int));
        table.Columns.Add("Content", typeof(string));
        table.Rows.Add(1, "<a href='https://example.com'>Example Link</a>");
        table.Rows.Add(2, "<b>Bold Text</b> and <i>Italic Text</i>");

        // Configure import options: do NOT preserve HTML tags
        // In Aspose.Cells the property that controls this behavior is IsHtmlString.
        // Setting it to false ensures that HTML tags are stripped from the imported values.
        ImportTableOptions importOptions = new ImportTableOptions
        {
            IsHtmlString = false,          // equivalent to PreserveHtml = false
            IsFieldNameShown = false       // we don't need column headers for this demo
        };

        // Import the DataTable into the worksheet starting at cell A1 (row 0, column 0)
        cells.ImportData(table, 0, 0, importOptions);

        // Verify that HTML tags have been removed from the imported cells
        for (int r = 0; r < table.Rows.Count; r++)
        {
            // Column index 1 corresponds to the "Content" column (B column in Excel)
            string cellValue = cells[r, 1].StringValue;
            bool containsHtmlTag = cellValue.Contains("<") && cellValue.Contains(">");
            Console.WriteLine($"Row {r + 1} - Cell B{r + 1}: \"{cellValue}\" | HTML tags removed: {!containsHtmlTag}");
        }

        // Save the workbook to verify the result manually if needed
        workbook.Save("ImportWithoutHtml.xlsx");
    }
}