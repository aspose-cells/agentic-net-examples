using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeUpdateDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook wb = new Workbook();                     // create
            Worksheet ws = wb.Worksheets[0];                 // first worksheet

            // ---------- Populate some data ----------
            ws.Cells["A1"].PutValue(10);
            ws.Cells["A2"].PutValue(20);
            ws.Cells["A3"].PutValue(30);

            // ---------- Define a named range that refers to A1:A3 ----------
            int nameIndex = wb.Worksheets.Names.Add("MyRange");
            Name myRange = wb.Worksheets.Names[nameIndex];
            myRange.RefersTo = "=Sheet1!$A$1:$A$3";

            // Use the named range in a formula (sum of the range)
            ws.Cells["B1"].Formula = "=SUM(MyRange)";

            // ---------- Initial calculation ----------
            wb.CalculateFormula();   // calculate formulas before any change

            Console.WriteLine("Initial SUM result: " + ws.Cells["B1"].Value); // should be 60

            // ---------- Update a cell inside the named range ----------
            ws.Cells["A2"].PutValue(50);   // change value from 20 to 50

            // ---------- Propagate changes ----------
            wb.CalculateFormula();   // recalculate dependent formulas

            Console.WriteLine("Updated SUM result: " + ws.Cells["B1"].Value); // should be 90

            // ---------- Save the workbook ----------
            wb.Save("NamedRangeUpdateDemo.xlsx");   // save
        }
    }
}