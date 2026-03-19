using System;
using Aspose.Cells;

namespace AsposeCellsDisplayRowsColumns
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data with different formatting
            // Row 0
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            // Apply bold font to header row
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            sheet.Cells["A1"].SetStyle(headerStyle);
            sheet.Cells["B1"].SetStyle(headerStyle);

            // Row 1
            sheet.Cells["A2"].PutValue("Data1");
            sheet.Cells["B2"].PutValue(123);
            // Apply background color to a data cell
            Style dataStyle = workbook.CreateStyle();
            dataStyle.ForegroundColor = System.Drawing.Color.LightYellow;
            dataStyle.Pattern = BackgroundType.Solid;
            sheet.Cells["B2"].SetStyle(dataStyle);

            // Row 2
            sheet.Cells["A3"].PutValue(DateTime.Now);
            sheet.Cells["B3"].PutValue(true);

            // Determine the used range
            int maxRow = sheet.Cells.MaxDataRow;      // last row with data
            int maxCol = sheet.Cells.MaxDataColumn;   // last column with data

            Console.WriteLine("Worksheet Cells Information:");
            Console.WriteLine("-----------------------------------");

            // Iterate through rows and columns within the used range
            for (int row = 0; row <= maxRow; row++)
            {
                // Access the Row object (Rows collection)
                var rowObj = sheet.Cells.Rows[row];
                Console.WriteLine($"Row {row + 1} (Height: {rowObj.Height} pts, Hidden: {rowObj.IsHidden})");

                for (int col = 0; col <= maxCol; col++)
                {
                    // Use GetCellDisplayStyle to obtain the effective style without instantiating empty cells
                    Style displayStyle = sheet.Cells.GetCellDisplayStyle(row, col);

                    // Retrieve cell value safely (CheckCell returns null if cell does not exist)
                    Cell cell = sheet.Cells.CheckCell(row, col);
                    string cellAddress = CellsHelper.CellIndexToName(row, col);
                    string value = cell != null ? cell.StringValue : string.Empty;

                    // Extract some formatting details
                    bool isBold = displayStyle.Font.IsBold;
                    System.Drawing.Color bgColor = displayStyle.ForegroundColor;
                    bool hasBackground = displayStyle.Pattern == BackgroundType.Solid;

                    Console.WriteLine($"  Cell {cellAddress}: Value=\"{value}\", Bold={isBold}, " +
                                      $"Background={(hasBackground ? bgColor.Name : "None")}");
                }
            }

            // Save the workbook (save rule)
            workbook.Save("RowsColumnsDisplay.xlsx");
        }
    }
}