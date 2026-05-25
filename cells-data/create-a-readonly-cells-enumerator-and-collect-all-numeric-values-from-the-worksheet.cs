using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsReadOnlyEnumerator
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (mixed types)
            cells["A1"].PutValue("Text");
            cells["B1"].PutValue(123);               // numeric
            cells["C1"].PutValue(45.67);              // numeric
            cells["D1"].PutValue(DateTime.Now);       // numeric (datetime)
            cells["E1"].PutValue(true);               // non‑numeric
            cells["F1"].PutValue("789");              // string that can be converted

            // Convert convertible strings to numeric values (optional)
            cells.ConvertStringToNumericValue();

            // Get a read‑only enumerator for all instantiated cells
            IEnumerator enumerator = cells.GetEnumerator();

            // List to hold numeric values found in the worksheet
            List<object> numericValues = new List<object>();

            // Iterate through cells without modifying the collection
            while (enumerator.MoveNext())
            {
                Cell cell = (Cell)enumerator.Current;
                // Check if the cell contains a numeric value (int, double, DateTime)
                if (cell.IsNumericValue && cell.Value != null)
                {
                    numericValues.Add(cell.Value);
                }
            }

            // Output the collected numeric values
            Console.WriteLine("Numeric values found in the worksheet:");
            foreach (var val in numericValues)
            {
                Console.WriteLine(val);
            }

            // Save the workbook (optional, demonstrates lifecycle rule usage)
            workbook.Save("ReadOnlyEnumeratorDemo.xlsx");
        }
    }
}