using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class UpdateTableColumnFormula
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Set up header row
        ws.Cells["A1"].PutValue("ID");
        ws.Cells["B1"].PutValue("Value");
        ws.Cells["C1"].PutValue("Result");

        // Add initial data rows
        ws.Cells["A2"].PutValue(1);
        ws.Cells["B2"].PutValue(10);
        ws.Cells["A3"].PutValue(2);
        ws.Cells["B3"].PutValue(20);
        ws.Cells["A4"].PutValue(3);
        ws.Cells["B4"].PutValue(30);

        // Create a table (list object) that includes the header and data
        int tableIndex = ws.ListObjects.Add("A1", "C4", true);
        ListObject table = ws.ListObjects[tableIndex];

        // Access the "Result" column (third column, index 2) and set an IF formula
        // The formula uses a structured reference to the current row's Value column
        ListColumn resultColumn = table.ListColumns[2];
        resultColumn.Formula = "=IF([@Value]>15,\"High\",\"Low\")";

        // Add new rows; the column formula will be applied automatically
        // Row offsets are zero‑based relative to the first data row (row after the header)
        table.PutCellValue(3, 0, 4);   // ID = 4
        table.PutCellValue(3, 1, 12);  // Value = 12 → Result should be "Low"
        table.PutCellValue(4, 0, 5);   // ID = 5
        table.PutCellValue(4, 1, 25);  // Value = 25 → Result should be "High"

        // Recalculate all formulas in the workbook
        wb.CalculateFormula();

        // Verify the Result column values by printing them to the console
        Console.WriteLine("Table data after adding rows:");
        for (int i = 0; i < table.DataRange.RowCount; i++)
        {
            int row = table.DataRange.FirstRow + i;
            string id = ws.Cells[row, 0].StringValue;
            string val = ws.Cells[row, 1].StringValue;
            string res = ws.Cells[row, 2].StringValue;
            Console.WriteLine($"ID={id}, Value={val}, Result={res}");
        }

        // Save the workbook
        wb.Save("TableIfFormulaDemo.xlsx");
    }
}