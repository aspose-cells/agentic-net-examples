// Title: Filter string‑type cells with Aspose.Cells for .NET (C#) and collect them into a List
// Description: Creates a workbook, fills cells with mixed data, enumerates every cell using the Cells iterator, selects only those whose Type equals CellValueType.IsString, adds the matching Cell objects to a List<Cell>, prints each address and string value, and optionally saves the file.
// Keywords: Aspose.Cells C# filter string cells | CellValueType.IsString example | enumerate worksheet cells .NET | collect string cells list | Aspose.Cells cell enumeration
// Common Searches: Aspose.Cells filter only string cells | C# enumerate all cells and get text values | How to collect string cells into a List with Aspose.Cells | CellValueType.IsString usage in Aspose.Cells
// Developer Intent: Iterate through every cell in a worksheet and build a List<Cell> containing only cells whose value type is string.
// Use Cases: Extract all textual entries for reporting or analytics. | Apply custom formatting or validation to text‑only cells. | Transfer string cells to another workbook or external system.
// AI Prompts: Generate C# code that uses Aspose.Cells to enumerate all worksheet cells and returns a List<Cell> of cells where Cell.Type is CellValueType.IsString. | Show an example that filters cells by CellValueType.IsString, prints each cell's address and string value, and saves the workbook. | Explain how to extend the enumeration to include cells whose evaluated formula result is a string while still using Aspose.Cells.

using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

// Creates a workbook, fills cells with mixed data, enumerates every cell using the Cells iterator, selects only those whose Type equals CellValueType.IsString, adds the matching Cell objects to a List<Cell>, prints each address and string value, and optionally saves the file.
class FilterStringCells
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample data: mix of strings and other types
        cells["A1"].PutValue("Hello");
        cells["B1"].PutValue(123);
        cells["A2"].PutValue("World");
        cells["B2"].PutValue(DateTime.Now);
        cells["C1"].PutValue("Aspose");
        cells["C2"].PutValue(true);

        // List to collect cells whose value type is string
        List<Cell> stringCells = new List<Cell>();

        // Enumerate all cells in the worksheet
        IEnumerator enumerator = cells.GetEnumerator();
        while (enumerator.MoveNext())
        {
            Cell cell = (Cell)enumerator.Current;

            // Filter: keep only cells with string value type
            if (cell.Type == CellValueType.IsString)
            {
                stringCells.Add(cell);
            }
        }

        // Display the addresses and string values of the collected cells
        Console.WriteLine("String cells found:");
        foreach (Cell sc in stringCells)
        {
            Console.WriteLine($"{sc.Name}: {sc.StringValue}");
        }

        // Save the workbook (optional)
        workbook.Save("FilteredStringCells.xlsx");
    }
}
