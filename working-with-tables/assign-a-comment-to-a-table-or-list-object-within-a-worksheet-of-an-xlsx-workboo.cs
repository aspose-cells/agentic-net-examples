using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsCommentOnTable
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data for the table
            worksheet.Cells["A1"].PutValue("ID");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue(100);
            worksheet.Cells["A3"].PutValue(2);
            worksheet.Cells["B3"].PutValue(200);

            // Add a ListObject (table) covering the data range
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
            int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 1, true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Assign a comment to the table using the ListObject.Comment property
            table.Comment = "This table contains sample ID and Value data.";

            // Save the workbook (lifecycle: save)
            workbook.Save("TableWithComment.xlsx", SaveFormat.Xlsx);
        }
    }
}