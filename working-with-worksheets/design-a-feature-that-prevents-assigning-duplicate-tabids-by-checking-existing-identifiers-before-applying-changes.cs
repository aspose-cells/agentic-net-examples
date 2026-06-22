using System;
using System.Collections.Generic;
using Aspose.Cells;

class PreventDuplicateTabIdDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the default worksheet and add two more worksheets
        Worksheet ws1 = workbook.Worksheets[0];
        ws1.Name = "Sheet1";

        Worksheet ws2 = workbook.Worksheets.Add("Sheet2");
        Worksheet ws3 = workbook.Worksheets.Add("Sheet3");

        // Assign TabIds while ensuring uniqueness
        SetUniqueTabId(ws1, 101, workbook);
        SetUniqueTabId(ws2, 102, workbook);
        // Attempt to assign a duplicate TabId (101) to ws3; the method will resolve it
        SetUniqueTabId(ws3, 101, workbook);

        // Save the workbook
        workbook.Save("UniqueTabIds.xlsx");
    }

    // Assigns a TabId to a worksheet, adjusting it if the desired value already exists
    static void SetUniqueTabId(Worksheet sheet, int desiredId, Workbook wb)
    {
        // Gather all existing TabIds in the workbook
        HashSet<int> existingIds = new HashSet<int>();
        foreach (Worksheet ws in wb.Worksheets)
        {
            existingIds.Add(ws.TabId);
        }

        int finalId = desiredId;
        // Find the next available TabId if the desired one is already taken
        while (existingIds.Contains(finalId))
        {
            finalId++;
        }

        // Apply the unique TabId
        sheet.TabId = finalId;

        // Output the result
        if (finalId != desiredId)
        {
            Console.WriteLine($"Desired TabId {desiredId} for worksheet '{sheet.Name}' was already used. Assigned new TabId {finalId}.");
        }
        else
        {
            Console.WriteLine($"Assigned TabId {finalId} to worksheet '{sheet.Name}'.");
        }
    }
}