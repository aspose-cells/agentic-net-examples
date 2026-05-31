using Aspose.Cells;
using Aspose.Cells.Tables;

class TableToRangeHeaderAndFirstRowDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate header row (row 0)
        cells[0, 0].PutValue("Header1");
        cells[0, 1].PutValue("Header2");

        // Populate first data row (row 1)
        cells[1, 0].PutValue("Data1");
        cells[1, 1].PutValue(100);

        // Populate additional data rows (these will be excluded from conversion)
        cells[2, 0].PutValue("Data2");
        cells[2, 1].PutValue(200);
        cells[3, 0].PutValue("Data3");
        cells[3, 1].PutValue(300);

        // Add a table that spans rows 0‑3 and columns 0‑1
        int tableIndex = worksheet.ListObjects.Add(0, 0, 3, 1, true);
        ListObject table = worksheet.ListObjects[tableIndex];
        table.TableStyleType = TableStyleType.TableStyleMedium2; // optional styling

        // Configure conversion options to retain only the header and the first data row
        TableToRangeOptions options = new TableToRangeOptions
        {
            // LastRow is zero‑based; setting it to 1 keeps rows 0 (header) and 1 (first data row)
            LastRow = 1
        };

        // Convert the table to a range using the specified options
        table.ConvertToRange(options);

        // Save the workbook
        workbook.Save("TableToRangeHeaderAndFirstRow.xlsx");
    }
}