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

            // Populate sample data (numeric and non‑numeric)
            cells["A1"].PutValue(10);          // numeric
            cells["B1"].PutValue("Text");      // non‑numeric
            cells["C1"].PutValue(25.5);        // numeric
            cells["A2"].PutValue(DateTime.Now); // date (numeric type)
            cells["B2"].PutValue(true);        // non‑numeric
            cells["C2"].PutValue("123");       // string that can be converted, but still a string

            // Get a read‑only enumerator for all cells in the worksheet
            IEnumerator enumerator = cells.GetEnumerator();

            // List to collect numeric values
            List<double> numericValues = new List<double>();

            // Iterate through cells without modifying the collection
            while (enumerator.MoveNext())
            {
                Cell cell = (Cell)enumerator.Current;

                // Check if the cell contains a numeric value (int, double, DateTime, etc.)
                if (cell != null && cell.IsNumericValue)
                {
                    // For numeric cells, DoubleValue provides a double representation
                    numericValues.Add(cell.DoubleValue);
                }
            }

            // Output the collected numeric values
            Console.WriteLine("Numeric values found in the worksheet:");
            foreach (double val in numericValues)
            {
                Console.WriteLine(val);
            }

            // Save the workbook (optional, demonstrates lifecycle rule usage)
            workbook.Save("ReadOnlyEnumeratorDemo.xlsx");
        }
    }
}