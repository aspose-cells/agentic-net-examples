using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Ods;

namespace AsposeCellsTableToRangeOds
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (5 rows, 3 columns)
            // Header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            // Data rows
            for (int row = 2; row <= 5; row++)
            {
                sheet.Cells[row - 1, 0].PutValue(row - 1);                     // ID
                sheet.Cells[row - 1, 1].PutValue($"Person {row - 1}");        // Name
                sheet.Cells[row - 1, 2].PutValue((row - 1) * 10);             // Score
            }

            // Apply formatting to the first three rows (header + first two data rows)
            Style boldStyle = workbook.CreateStyle();
            boldStyle.Font.IsBold = true;
            StyleFlag flag = new StyleFlag();
            flag.FontBold = true;

            // Apply to rows 0,1,2 (zero‑based)
            sheet.Cells.Rows[0].ApplyStyle(boldStyle, flag);
            sheet.Cells.Rows[1].ApplyStyle(boldStyle, flag);
            sheet.Cells.Rows[2].ApplyStyle(boldStyle, flag);

            // Add a ListObject (table) that covers the whole data range A1:C5
            int tableIndex = sheet.ListObjects.Add("A1", "C5", true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Convert only the first three rows of the table to a normal range,
            // preserving the formatting applied above.
            TableToRangeOptions options = new TableToRangeOptions
            {
                // Row indices are zero‑based; 2 corresponds to the third row.
                LastRow = 2
            };
            table.ConvertToRange(options);

            // Save the workbook as ODS using default OdsSaveOptions
            OdsSaveOptions odsOptions = new OdsSaveOptions();
            workbook.Save("TableConvertedToRange.ods", odsOptions);
        }
    }
}