using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data with a header and a numeric column
        cells["A1"].PutValue("Item");
        cells["B1"].PutValue("Quantity");
        cells["A2"].PutValue("Apple");
        cells["B2"].PutValue(10);
        cells["A3"].PutValue("Banana");
        cells["B3"].PutValue(20);
        cells["A4"].PutValue("Cherry");
        cells["B4"].PutValue(15);

        // Create a table (ListObject) covering the data range A1:B4
        int tableIdx = worksheet.ListObjects.Add("A1", "B4", true);
        ListObject table = worksheet.ListObjects[tableIdx];

        // Retrieve the table using a cell that belongs to it
        Cell cellInTable = cells["B2"];
        ListObject retrievedTable = cellInTable.GetTable();

        // Determine the data rows (skip header) and the column index of the numeric data
        int dataStartRow = retrievedTable.StartRow + 1; // first data row after header
        int dataEndRow = retrievedTable.EndRow;         // last data row
        int quantityColumn = retrievedTable.StartColumn + 1; // second column (Quantity)

        // Compute the sum of the numeric column
        double sum = 0;
        for (int row = dataStartRow; row <= dataEndRow; row++)
        {
            object value = cells[row, quantityColumn].Value;
            if (value is double d)
                sum += d;
            else if (value is int i)
                sum += i;
            else if (double.TryParse(Convert.ToString(value), out double parsed))
                sum += parsed;
        }

        // Write the computed sum below the table
        cells[dataEndRow + 2, quantityColumn].PutValue(sum);

        // Save the workbook
        workbook.Save("TableSumResult.xlsx");
    }
}