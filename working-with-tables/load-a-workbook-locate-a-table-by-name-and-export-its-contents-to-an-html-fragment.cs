using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Tables;          // For ListObject
using AsposeRange = Aspose.Cells.Range; // Alias to avoid conflict with System.Range

class ExportTableToHtmlFragment
{
    static void Main()
    {
        try
        {
            // Path to the source Excel workbook
            string workbookPath = "input.xlsx";

            // Verify that the file exists before loading
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found: {workbookPath}");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(workbookPath);

            // Assume the table is in the first worksheet; adjust as needed
            Worksheet worksheet = workbook.Worksheets[0];

            // Locate the table (ListObject) by its name
            // Replace "MyTable" with the actual table name in the workbook
            ListObject table = worksheet.ListObjects["MyTable"];
            if (table == null)
            {
                Console.WriteLine("Table 'MyTable' not found in the worksheet.");
                return;
            }

            // Determine the cell area that contains the table data
            AsposeRange dataRange = table.DataRange;
            int startRow = dataRange.FirstRow;
            int startColumn = dataRange.FirstColumn;
            int endRow = startRow + dataRange.RowCount - 1;
            int endColumn = startColumn + dataRange.ColumnCount - 1;

            CellArea exportArea = CellArea.CreateCellArea(startRow, startColumn, endRow, endColumn);

            // Configure HTML save options to export only the table part
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportDataOptions = HtmlExportDataOptions.Table,
                ExportArea = exportArea
            };

            // Save the selected area to a memory stream as HTML
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, htmlOptions);
                ms.Position = 0;

                // Convert the HTML bytes to a string (HTML fragment)
                string htmlFragment = Encoding.UTF8.GetString(ms.ToArray());

                // Output the HTML fragment
                Console.WriteLine(htmlFragment);
            }
        }
        catch (Exception ex)
        {
            // Runtime safety: report any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}