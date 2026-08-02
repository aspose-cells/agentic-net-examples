// Title: C# – Iterate the Used Range in Aspose.Cells and Log Cell Addresses with Values
// Description: Demonstrates how to create a workbook, populate sample data, obtain an enumerator for the worksheet's used cells, and loop through each non‑empty cell to print its address (Name) and value to the console before saving the file.
// Keywords: Aspose.Cells used range iteration C# | enumerate worksheet cells Aspose | log cell address Aspose.Cells | skip empty cells Aspose | C# Aspose.Cells console output
// Common Searches: how to loop through used cells in Aspose.Cells C# | Aspose.Cells enumerate non‑empty cells | print cell address and value with Aspose.Cells | C# Aspose.Cells iterate used range example
// Developer Intent: Retrieve every populated cell in a worksheet and output its reference and content programmatically.
// Use Cases: Generate a quick console dump of all data entries for debugging. | Create an audit trail of worksheet contents by logging cell references. | Export non‑empty cell values to a custom report or file.
// AI Prompts: Write C# code using Aspose.Cells that iterates only non‑empty cells in the used range and writes each cell's address and formatted value to a text file. | Show how to modify the enumeration to include cell style information (e.g., font color) while iterating the used range. | Explain how to filter the iteration to process only numeric cells in Aspose.Cells C#.

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, populate sample data, obtain an enumerator for the worksheet's used cells, and loop through each non‑empty cell to print its address (Name) and value to the console before saving the file.
    class IterateUsedRange
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample data (optional, demonstrates the iteration)
            cells["A1"].PutValue("Header");
            cells["B1"].PutValue("Value");
            cells["A2"].PutValue(123);
            cells["B2"].PutValue(DateTime.Now);
            cells["C3"].PutValue("Extra");

            // Get an enumerator for all cells that contain data in the worksheet
            IEnumerator enumerator = cells.GetEnumerator();

            // Iterate through the cells and log address (Name) and value
            while (enumerator.MoveNext())
            {
                Cell cell = (Cell)enumerator.Current;
                // Some cells may be empty; skip if no value
                if (cell.Value != null)
                {
                    Console.WriteLine($"{cell.Name}: {cell.Value}");
                }
            }

            // Save the workbook (optional, just to persist any changes)
            workbook.Save("IterateUsedRange.xlsx");
        }
    }
}
