using System;
using System.Data;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ------------------------------------------------------------
            // 1. Create a source range using integer coordinates (A3:D5)
            // ------------------------------------------------------------
            // Parameters: firstRow, firstColumn, totalRows, totalColumns
            AsposeRange sourceRange = cells.CreateRange(2, 0, 3, 4); // rows 3-5, columns A-D

            // Fill the source range with sample data
            for (int i = 0; i < sourceRange.RowCount; i++)
            {
                for (int j = 0; j < sourceRange.ColumnCount; j++)
                {
                    sourceRange[i, j].PutValue($"R{i + 3}C{j + 1}");
                }
            }

            // Name the source range for later formula use
            sourceRange.Name = "MyData";

            // ------------------------------------------------------------
            // 2. Create a destination range using address strings (A7:D9)
            // ------------------------------------------------------------
            AsposeRange destRange = cells.CreateRange("A7", "D9");

            // Copy only the values from source to destination
            destRange.CopyValue(sourceRange);

            // ------------------------------------------------------------
            // 3. Use the named range in a formula
            // ------------------------------------------------------------
            cells["E1"].Formula = "=SUM(MyData)";
            workbook.CalculateFormula();

            // ------------------------------------------------------------
            // 4. Add the source range to the worksheet's RangeCollection
            // ------------------------------------------------------------
            cells.AddRange(sourceRange);

            // Insert rows to demonstrate that the range expands automatically
            // Insert 2 rows after row index 1 (i.e., before row 2)
            cells.InsertRows(1, 2, true);

            // ------------------------------------------------------------
            // 5. Export the (now expanded) source range to a DataTable
            // ------------------------------------------------------------
            DataTable dt = sourceRange.ExportDataTable();

            // Display exported data in the console
            Console.WriteLine("Exported DataTable contents:");
            foreach (DataRow row in dt.Rows)
            {
                foreach (var item in row.ItemArray)
                {
                    Console.Write(item + "\t");
                }
                Console.WriteLine();
            }

            // ------------------------------------------------------------
            // 6. Save the workbook to an XLSX file
            // ------------------------------------------------------------
            workbook.Save("ManagedRangesDemo.xlsx");

            Console.WriteLine("Workbook saved as ManagedRangesDemo.xlsx");
        }
    }
}