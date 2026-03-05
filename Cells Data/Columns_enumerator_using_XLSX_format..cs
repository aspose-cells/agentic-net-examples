using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data in the first three columns
        cells["A1"].PutValue("Header A");
        cells["B1"].PutValue("Header B");
        cells["C1"].PutValue("Header C");
        cells["A2"].PutValue(1);
        cells["B2"].PutValue(2);
        cells["C2"].PutValue(3);

        // Enumerate all instantiated columns in the worksheet
        Console.WriteLine("Enumerating columns:");
        foreach (Column col in sheet.Cells.Columns)
        {
            // Column index (zero‑based)
            int index = col.Index;

            // Convert index to Excel column name (A, B, C, …)
            string name = CellsHelper.ColumnIndexToName(index);

            // Retrieve the current width of the column
            double width = col.Width;

            Console.WriteLine($"Column {name} (Index {index}) - Width: {width}");
        }

        // Save the workbook in XLSX format
        workbook.Save("ColumnsEnumerator.xlsx");
    }
}