using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsStringFilterDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (mix of strings, numbers, dates, booleans)
            cells["A1"].PutValue("Apple");
            cells["B1"].PutValue(123);
            cells["C1"].PutValue(DateTime.Now);
            cells["A2"].PutValue("Banana");
            cells["B2"].PutValue(true);
            cells["C2"].PutValue("Cherry");

            // List to hold string‑type cell values
            List<string> stringValues = new List<string>();

            // Get the enumerator for all cells in the worksheet
            IEnumerator enumerator = cells.GetEnumerator();

            // Iterate through each cell
            while (enumerator.MoveNext())
            {
                Cell cell = (Cell)enumerator.Current;

                // Check if the cell's value type is string
                if (cell.Type == CellValueType.IsString)
                {
                    // Add the string value to the list
                    stringValues.Add(cell.StringValue);
                }
            }

            // Output the collected string values
            Console.WriteLine("String‑type cells collected:");
            foreach (string val in stringValues)
            {
                Console.WriteLine(val);
            }

            // Save the workbook (optional, demonstrates usage of save rule)
            workbook.Save("StringFilterResult.xlsx");
        }
    }
}