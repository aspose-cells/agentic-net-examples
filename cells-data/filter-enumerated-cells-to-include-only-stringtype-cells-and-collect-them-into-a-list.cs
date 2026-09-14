// Title: How to enumerate cells in an Aspose.Cells worksheet and collect only string‑type values into a List<string> (C#)
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, iterates through every cell in the first worksheet using the Cells enumerator, checks if the cell's Type equals CellValueType.IsString, and adds the cell's StringValue to a List<string>. | Show a C# example that filters out non‑text cells while enumerating a worksheet's cells with Aspose.Cells and returns a List<string> containing only the text values. | Demonstrate how to use Aspose.Cells in .NET to collect all string cells from a worksheet, print each value to the console, and then save the workbook without modifications.
// Common Searches: Aspose.Cells C# get only text cells from a worksheet | enumerate Excel cells and extract string values using Aspose.Cells | list of string values from Excel file with Aspose.Cells .NET | how to retrieve only string cells from a workbook in C# | Aspose.Cells filter non‑numeric cells example
// Tags: Aspose.Cells enumerate worksheet cells C# | filter cells by CellValueType.IsString | extract string values from Excel worksheet | collect text cells into List<string> | Aspose.Cells read-only workbook processing

using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

// The program loads an Excel workbook with Aspose.Cells, enumerates every cell in the first worksheet, adds the StringValue of cells whose Type is CellValueType.IsString to a List<string>, prints the collected strings, and saves the workbook unchanged.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace the path with your actual file)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // List to store string values from cells
        List<string> stringValues = new List<string>();

        // Enumerate all cells in the worksheet
        IEnumerator enumerator = cells.GetEnumerator();
        while (enumerator.MoveNext())
        {
            Cell cell = (Cell)enumerator.Current;

            // Filter only cells whose type is string
            if (cell.Type == CellValueType.IsString)
            {
                stringValues.Add(cell.StringValue);
            }
        }

        // Display the collected string values
        Console.WriteLine("Collected string cells:");
        foreach (string value in stringValues)
        {
            Console.WriteLine(value);
        }

        // Save the workbook (no modifications made, just demonstrating save lifecycle)
        workbook.Save("output.xlsx");
    }
}
