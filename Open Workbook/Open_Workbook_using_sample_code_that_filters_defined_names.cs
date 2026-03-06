using Aspose.Cells;
using System;

class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string sourceFile = "input.xlsx";

        // Load the workbook from the file
        Workbook workbook = new Workbook(sourceFile);

        // Get defined names that have workbook scope
        Name[] workbookScopeNames = workbook.Worksheets.Names.Filter(NameScopeType.Workbook, -1);
        Console.WriteLine("Workbook scoped names count: " + workbookScopeNames.Length);

        // Get all defined names that have worksheet scope (across all sheets)
        Name[] allWorksheetScopeNames = workbook.Worksheets.Names.Filter(NameScopeType.Worksheet, -1);
        Console.WriteLine("All worksheet scoped names count: " + allWorksheetScopeNames.Length);

        // Get defined names for a specific worksheet (e.g., first sheet, index 0)
        Name[] sheet0Names = workbook.Worksheets.Names.Filter(NameScopeType.Worksheet, 0);
        Console.WriteLine("Sheet 0 scoped names count: " + sheet0Names.Length);

        // Save the workbook (no modifications made, just demonstrating save)
        workbook.Save("FilteredNamesResult.xlsx");
    }
}