using System;
using Aspose.Cells;

namespace AsposeCellsCellManipulationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the Cells collection from the worksheet
            Cells cells = worksheet.Cells;

            // -------------------------------------------------
            // Manipulate individual cells using row/column indexes
            // -------------------------------------------------
            // Set value in cell A1 (row 0, column 0)
            cells[0, 0].PutValue("Hello");

            // Set numeric value in cell C3 (row 2, column 2)
            cells[2, 2].PutValue(3.14159);

            // -------------------------------------------------
            // Manipulate individual cells using cell names
            // -------------------------------------------------
            // Set a string value in cell B2
            cells["B2"].PutValue("World");

            // Set an integer value in cell D4
            cells["D4"].PutValue(42);

            // -------------------------------------------------
            // Retrieve values from cells
            // -------------------------------------------------
            string valueA1 = cells["A1"].StringValue;          // "Hello"
            string valueB2 = cells["B2"].StringValue;          // "World"
            double valueC3 = cells[2, 2].DoubleValue;          // 3.14159
            int valueD4 = cells["D4"].IntValue;                // 42

            // Output retrieved values to the console
            Console.WriteLine($"A1: {valueA1}");
            Console.WriteLine($"B2: {valueB2}");
            Console.WriteLine($"C3: {valueC3}");
            Console.WriteLine($"D4: {valueD4}");

            // Save the workbook to disk (lifecycle: save)
            workbook.Save("ManipulatedCells.xlsx");

            // Dispose the workbook to release resources
            workbook.Dispose();
        }
    }
}