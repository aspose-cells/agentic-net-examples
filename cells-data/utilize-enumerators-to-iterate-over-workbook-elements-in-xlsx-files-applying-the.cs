using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsEnumeratorsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default Xlsx format)
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "FirstSheet";
            sheet1.Cells["A1"].PutValue("Name");
            sheet1.Cells["B1"].PutValue("Age");
            sheet1.Cells["A2"].PutValue("John");
            sheet1.Cells["B2"].PutValue(30);
            sheet1.Cells["A3"].PutValue("Alice");
            sheet1.Cells["B3"].PutValue(25);

            // Add a second worksheet and populate it
            int sheet2Index = workbook.Worksheets.Add();
            Worksheet sheet2 = workbook.Worksheets[sheet2Index];
            sheet2.Name = "SecondSheet";
            sheet2.Cells["A1"].PutValue("Product");
            sheet2.Cells["B1"].PutValue("Price");
            sheet2.Cells["A2"].PutValue("Apple");
            sheet2.Cells["B2"].PutValue(1.2);
            sheet2.Cells["A3"].PutValue("Banana");
            sheet2.Cells["B3"].PutValue(0.8);

            // Iterate over all worksheets using the worksheet collection enumerator
            IEnumerator sheetEnumerator = workbook.Worksheets.GetEnumerator();
            while (sheetEnumerator.MoveNext())
            {
                Worksheet ws = (Worksheet)sheetEnumerator.Current;
                Console.WriteLine($"--- Worksheet: {ws.Name} ---");

                // Iterate over rows in the current worksheet
                IEnumerator rowEnumerator = ws.Cells.Rows.GetEnumerator();
                while (rowEnumerator.MoveNext())
                {
                    Row row = (Row)rowEnumerator.Current;

                    // Iterate over cells in the current row
                    IEnumerator cellEnumerator = row.GetEnumerator();
                    while (cellEnumerator.MoveNext())
                    {
                        Cell cell = (Cell)cellEnumerator.Current;
                        if (cell.Value != null)
                        {
                            Console.WriteLine($"{cell.Name}: {cell.Value}");
                        }
                    }
                }
            }

            // Save the workbook to disk
            workbook.Save("EnumeratorsDemo.xlsx");
        }
    }
}