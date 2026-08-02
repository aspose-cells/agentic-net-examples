// Title: Detect Null or Empty Cells While Enumerating Worksheet Cells – Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, inserts intentional blank cells, iterates through all worksheet cells with Cells.GetEnumerator(), checks each Cell.Value for null, logs the address of every empty cell, and saves the file. Ideal for data‑completeness checks in Excel automation.
// Keywords: Aspose.Cells C# null cell detection | enumerate worksheet cells .NET | detect empty Excel cells Aspose | Cell.Value null check | data completeness Aspose.Cells | Excel blank cell logging | GetEnumerator example | Aspose.Cells sample code | C# Excel data validation | Aspose.Cells API usage
// Common Searches: how to find blank cells with Aspose.Cells | enumerate cells and detect null values in C# | Aspose.Cells log empty cells during iteration | detect missing data in Excel using Aspose.Cells | C# code to list null cells in a worksheet
// Developer Intent: Identify every cell that contains a null (or missing) value while iterating through a worksheet and record its address.
// Use Cases: Audit imported spreadsheets for missing entries before processing. | Generate a report of all empty cells to improve data quality. | Skip or replace null cells in bulk calculations to avoid runtime errors.
// AI Prompts: Write C# code using Aspose.Cells that collects the addresses of all null cells into a List<string>. | Show how to extend the enumeration loop to treat empty strings and whitespace as null‑like values and log them. | Provide an example that writes detected null cell addresses to a text file instead of the console.

using System;
using System.Collections;
using Aspose.Cells;

// C# example that creates a workbook, inserts intentional blank cells, iterates through all worksheet cells with Cells.GetEnumerator(), checks each Cell.Value for null, logs the address of every empty cell, and saves the file. Ideal for data‑completeness checks in Excel automation.
class DetectNullCells
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Add sample data with intentional empty (null) cells
        cells["A1"].PutValue("Header1");
        cells["B1"].PutValue("Header2");
        cells["A2"].PutValue(100);
        // B2 left empty – will be null
        cells["A3"].PutValue(null); // explicit null assignment
        cells["B3"].PutValue(300);

        // Enumerate all cells in the worksheet
        IEnumerator enumerator = cells.GetEnumerator();
        while (enumerator.MoveNext())
        {
            Cell cell = (Cell)enumerator.Current;

            // Log cells whose value is null
            if (cell.Value == null)
            {
                Console.WriteLine($"Null cell detected at {cell.Name}");
            }
        }

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("NullCellDetection.xlsx");
    }
}
