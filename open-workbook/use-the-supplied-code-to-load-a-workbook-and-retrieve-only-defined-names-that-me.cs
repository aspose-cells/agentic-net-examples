using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Get all defined names that have workbook scope
        Name[] workbookScopeNames = workbook.Worksheets.Names.Filter(NameScopeType.Workbook, -1);
        Console.WriteLine($"Workbook-scoped names count: {workbookScopeNames.Length}");
        foreach (Name name in workbookScopeNames)
        {
            Console.WriteLine($"Name: {name.Text}, RefersTo: {name.RefersTo}");
        }

        // Get all defined names that have worksheet scope (across all worksheets)
        Name[] worksheetScopeNames = workbook.Worksheets.Names.Filter(NameScopeType.Worksheet, -1);
        Console.WriteLine($"Worksheet-scoped names count (all sheets): {worksheetScopeNames.Length}");

        // Get defined names for a specific worksheet (e.g., first sheet, index 0)
        int sheetIndex = 0;
        Name[] sheetNames = workbook.Worksheets.Names.Filter(NameScopeType.Worksheet, sheetIndex);
        Console.WriteLine($"Names in sheet index {sheetIndex} count: {sheetNames.Length}");
        foreach (Name name in sheetNames)
        {
            Console.WriteLine($"Name: {name.Text}, RefersTo: {name.RefersTo}");
        }

        // Save the workbook if any changes were made (not required in this example)
        // workbook.Save("output.xlsx");
    }
}