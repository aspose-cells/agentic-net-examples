// Title: Extract only string‑type cells from an Excel worksheet into a List<string> using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that scans every cell in a worksheet and appends only cells whose Type equals IsString to a List<string>. | Adapt an Aspose.Cells example to skip numeric, date, and formula cells and return a collection of the cells' StringValue properties.
// Common Searches: Aspose.Cells C# get text values from all cells in a worksheet | How to filter Excel cells by string type using Aspose.Cells .NET | Collect only string cells into a List in C# with Aspose.Cells | Iterate worksheet cells and retrieve only string values Aspose.Cells example | C# Aspose.Cells enumerate cells and extract only string entries
// Tags: filter string cells Aspose.Cells | extract text values from Excel worksheet .NET | enumerate worksheet cells Aspose.Cells C# | list<string> of Excel string cells Aspose.Cells | cell type checking IsString Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, iterates through every cell in the first worksheet, checks each cell's Type for IsString, adds the StringValue of matching cells to a List<string>, and outputs the collected strings.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (or specify the desired one)
        Worksheet sheet = workbook.Worksheets[0];

        // List to hold string‑type cell values
        List<string> stringValues = new List<string>();

        // Iterate through all cells in the worksheet
        foreach (Cell cell in sheet.Cells)
        {
            // Check if the cell contains a string value
            if (cell.Type == CellValueType.IsString)
            {
                // Add the string value to the list
                stringValues.Add(cell.StringValue);
            }
        }

        // Example usage: print collected strings
        foreach (string s in stringValues)
        {
            Console.WriteLine(s);
        }
    }
}
