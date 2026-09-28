// Title: Count how many worksheet cells contain a specific keyword using Aspose.Cells enumeration in C#
// AI Prompts: Generate C# code that opens a workbook, enumerates every cell with Aspose.Cells, and returns the count of cells whose text includes a supplied keyword. | Adapt the enumeration loop to perform a case‑insensitive keyword search and optionally capture the addresses of matching cells.
// Common Searches: how to count cells containing a word with Aspose.Cells in C# | Aspose.Cells enumerate worksheet cells and find substring | C# count cells that include "Apple" using Aspose.Cells | search for keyword in Excel cells with Aspose.Cells enumeration | case‑insensitive keyword count in an Aspose.Cells worksheet
// Tags: enumerate cells Aspose.Cells C# | keyword search in Excel worksheet Aspose.Cells | count matching cells Aspose.Cells | cell string value contains check C# | case‑sensitive keyword count Aspose.Cells

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsKeywordCount
{
    // The example creates a workbook, fills several cells with sample text, then uses Aspose.Cells' cell enumerator to walk through all instantiated cells. For each cell it checks whether the StringValue contains the target keyword (case‑sensitive) and increments a counter. The final count is printed and the workbook is saved as KeywordCountDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue("Apple");
            worksheet.Cells["B1"].PutValue("Banana");
            worksheet.Cells["A2"].PutValue("Green Apple");
            worksheet.Cells["B2"].PutValue("Orange");
            worksheet.Cells["C3"].PutValue("Pineapple");
            worksheet.Cells["D4"].PutValue("Apple Pie");

            // Keyword to search for
            string keyword = "Apple";

            // Counter for cells containing the keyword
            int keywordCount = 0;

            // Get the cells enumerator for the worksheet
            IEnumerator cellEnumerator = worksheet.Cells.GetEnumerator();

            // Iterate through all instantiated cells
            while (cellEnumerator.MoveNext())
            {
                Cell cell = (Cell)cellEnumerator.Current;

                // Ensure the cell has a string representation
                if (cell != null && cell.StringValue != null)
                {
                    // Check if the cell's text contains the keyword (case‑sensitive)
                    if (cell.StringValue.Contains(keyword))
                    {
                        keywordCount++;
                    }
                }
            }

            // Output the result
            Console.WriteLine($"Number of cells containing \"{keyword}\": {keywordCount}");

            // Save the workbook (optional, just to demonstrate lifecycle compliance)
            workbook.Save("KeywordCountDemo.xlsx");
        }
    }
}
