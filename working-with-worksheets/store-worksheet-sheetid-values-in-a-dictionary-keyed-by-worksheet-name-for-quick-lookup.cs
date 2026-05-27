using System;
using System.Collections.Generic;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Add sample worksheets for demonstration
        workbook.Worksheets[0].Name = "First";
        workbook.Worksheets.Add("Second");
        workbook.Worksheets.Add("Third");

        // Dictionary to map worksheet name -> internal SheetId (TabId)
        Dictionary<string, int> sheetIdLookup = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        // Populate the dictionary with each worksheet's Name and TabId
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheetIdLookup[sheet.Name] = sheet.TabId;
        }

        // Example: retrieve the TabId for a given sheet name
        string targetSheet = "Second";
        if (sheetIdLookup.TryGetValue(targetSheet, out int tabId))
        {
            Console.WriteLine($"Sheet '{targetSheet}' has TabId: {tabId}");
        }
        else
        {
            Console.WriteLine($"Sheet '{targetSheet}' not found in the workbook.");
        }

        // Save the workbook (optional)
        workbook.Save("Output.xlsx");
    }
}