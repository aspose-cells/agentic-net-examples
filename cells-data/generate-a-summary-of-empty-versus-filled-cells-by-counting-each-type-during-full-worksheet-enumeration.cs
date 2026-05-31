using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

class EmptyVsFilledSummary
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Add some sample data (the rest of the cells remain empty)
        cells["A1"].PutValue("Sample Text");
        cells["B2"].PutValue(12345);
        cells["C3"].PutValue(DateTime.Now);
        cells["D4"].PutValue(true);

        // Dictionary to store counts for each CellValueType
        Dictionary<CellValueType, long> typeCounts = new Dictionary<CellValueType, long>();

        // Enumerate all instantiated cells in the worksheet
        IEnumerator enumerator = cells.GetEnumerator();
        while (enumerator.MoveNext())
        {
            Cell cell = (Cell)enumerator.Current;
            CellValueType type = cell.Type; // Determines the cell's data type

            if (!typeCounts.ContainsKey(type))
                typeCounts[type] = 0;
            typeCounts[type]++;
        }

        // Total number of instantiated cells (including empty ones)
        long totalInstantiated = cells.CountLarge;

        // Output the count for each cell type
        Console.WriteLine($"Total instantiated cells: {totalInstantiated}");
        foreach (var kvp in typeCounts)
        {
            Console.WriteLine($"{kvp.Key}: {kvp.Value}");
        }

        // Separate empty (IsNull) and filled cells
        long emptyCount = typeCounts.ContainsKey(CellValueType.IsNull) ? typeCounts[CellValueType.IsNull] : 0;
        long filledCount = totalInstantiated - emptyCount;

        Console.WriteLine($"Empty cells: {emptyCount}");
        Console.WriteLine($"Filled cells: {filledCount}");

        // Save the workbook (demonstrates the required save lifecycle)
        workbook.Save("Summary.xlsx", SaveFormat.Xlsx);
    }
}