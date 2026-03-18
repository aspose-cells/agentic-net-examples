using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class EnumerateWorkbookDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Populate sample data in the first worksheet
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "FirstSheet";
            sheet1.Cells["A1"].PutValue("Sheet1-A1");
            sheet1.Cells["B2"].PutValue(123);
            sheet1.Cells["C3"].PutValue(DateTime.Now);

            // Add a second worksheet and fill it with data
            int sheet2Index = workbook.Worksheets.Add();
            Worksheet sheet2 = workbook.Worksheets[sheet2Index];
            sheet2.Name = "SecondSheet";
            sheet2.Cells["A1"].PutValue("Sheet2-A1");
            sheet2.Cells["A2"].PutValue("Sheet2-A2");
            sheet2.Cells["B1"].PutValue(456);

            // Iterate through worksheets using the enumerator provided by WorksheetCollection
            IEnumerator wsEnum = workbook.Worksheets.GetEnumerator();
            while (wsEnum.MoveNext())
            {
                Worksheet ws = (Worksheet)wsEnum.Current;
                Console.WriteLine($"Worksheet: {ws.Name}");

                // Iterate through rows in the current worksheet using RowCollection enumerator
                IEnumerator rowEnum = ws.Cells.Rows.GetEnumerator();
                while (rowEnum.MoveNext())
                {
                    Row row = (Row)rowEnum.Current;
                    Console.WriteLine($"  Row index: {row.Index}");

                    // Iterate through cells in the current row using Row enumerator
                    IEnumerator cellEnum = row.GetEnumerator();
                    while (cellEnum.MoveNext())
                    {
                        Cell cell = (Cell)cellEnum.Current;
                        if (cell.Value != null)
                        {
                            Console.WriteLine($"    Cell {cell.Name}: {cell.Value}");
                        }
                    }
                }
            }

            // Save the workbook
            workbook.Save("EnumerateWorkbookDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            EnumerateWorkbookDemo.Run();
        }
    }
}