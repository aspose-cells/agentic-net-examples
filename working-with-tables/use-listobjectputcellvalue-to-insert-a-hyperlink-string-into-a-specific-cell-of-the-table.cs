using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsHyperlinkInTable
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data for the table (including header)
            worksheet.Cells["A1"].PutValue("ID");
            worksheet.Cells["B1"].PutValue("Name");
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue("Alice");
            worksheet.Cells["A3"].PutValue(2);
            worksheet.Cells["B3"].PutValue("Bob");

            // Create a ListObject (table) that covers the data range A1:B3
            // Parameters: first row, first column, total rows, total columns, show header
            int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 1, true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Define the row and column offset inside the table where the hyperlink will be placed
            // For example, put hyperlink in the second data row (row offset 2) and second column (column offset 1)
            int rowOffset = 2;      // 0‑based offset from the start of the table (including header)
            int columnOffset = 1;   // second column (Name)

            // Insert the display text for the hyperlink using PutCellValue
            table.PutCellValue(rowOffset, columnOffset, "Visit Aspose");

            // Determine the absolute cell address (e.g., "B3") where the value was placed
            int absoluteRow = table.StartRow + rowOffset;
            int absoluteColumn = table.StartColumn + columnOffset;
            string cellName = worksheet.Cells[absoluteRow, absoluteColumn].Name; // e.g., "B3"

            // Add a hyperlink to the same cell
            // Parameters: cell name, total rows, total columns, hyperlink address
            worksheet.Hyperlinks.Add(cellName, 1, 1, "https://www.aspose.com");

            // Optionally, customize the displayed text (already set by PutCellValue)
            // worksheet.Hyperlinks[worksheet.Hyperlinks.Count - 1].TextToDisplay = "Visit Aspose";

            // Save the workbook
            workbook.Save("TableWithHyperlink.xlsx", SaveFormat.Xlsx);
        }
    }
}