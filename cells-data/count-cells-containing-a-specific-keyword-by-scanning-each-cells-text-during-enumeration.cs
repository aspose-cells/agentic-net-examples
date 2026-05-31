using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsKeywordCount
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue("Apple");
            worksheet.Cells["B1"].PutValue("Banana");
            worksheet.Cells["A2"].PutValue("Green Apple");
            worksheet.Cells["B2"].PutValue("Orange");
            worksheet.Cells["A3"].PutValue("Pineapple");
            worksheet.Cells["B3"].PutValue("Grape");

            // Keyword to search for
            string keyword = "Apple";

            // Counter for cells containing the keyword
            int count = 0;

            // Enumerate all instantiated cells (rule: Cells.GetEnumerator)
            IEnumerator enumerator = worksheet.Cells.GetEnumerator();
            while (enumerator.MoveNext())
            {
                Cell cell = (Cell)enumerator.Current;
                // Ensure the cell has a string representation
                if (cell != null && cell.StringValue != null && cell.StringValue.Contains(keyword))
                {
                    count++;
                }
            }

            Console.WriteLine($"Number of cells containing \"{keyword}\": {count}");

            // Save the workbook (lifecycle rule: save)
            workbook.Save("KeywordCountResult.xlsx");
        }
    }
}