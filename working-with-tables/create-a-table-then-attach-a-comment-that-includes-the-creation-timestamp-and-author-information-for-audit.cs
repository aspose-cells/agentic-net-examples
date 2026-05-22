using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableCommentDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Sample data for the table (header + two rows)
            worksheet.Cells["A1"].PutValue("ID");
            worksheet.Cells["B1"].PutValue("Name");
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue("Alice");
            worksheet.Cells["A3"].PutValue(2);
            worksheet.Cells["B3"].PutValue("Bob");

            // Define author information
            string author = "John Doe";

            // Create a ListObject (table) that covers the data range
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
            int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 1, true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Attach an audit comment to the table with timestamp and author
            table.Comment = $"Created by {author} on {DateTime.Now:yyyy-MM-dd HH:mm:ss}";

            // Save the workbook
            workbook.Save("TableWithAuditComment.xlsx", SaveFormat.Xlsx);
        }
    }
}