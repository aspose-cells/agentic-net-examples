// Title: Check that worksheet TabId values stay the same after saving and reloading a workbook using Aspose.Cells for .NET
// AI Prompts: Generate a C# routine that records each Worksheet.TabId, saves the workbook to a MemoryStream in XLSX format, reloads it, and reports any TabId mismatches. | Create an NUnit/xUnit test that asserts all Worksheet.TabId values are identical before and after workbook serialization with Aspose.Cells.
// Common Searches: Aspose.Cells how to verify worksheet TabId remains unchanged after saving | C# compare Excel sheet TabId values before and after workbook serialization | ensure tab identifier persistence when loading a workbook from a memory stream in Aspose.Cells | detect TabId mismatches in reloaded worksheets using Aspose.Cells .NET
// Tags: worksheet TabId persistence after save | Aspose.Cells workbook round‑trip TabId verification | C# compare worksheet TabId pre‑save post‑load | Excel sheet identifier consistency with Aspose.Cells | memory stream serialization TabId check

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example records each worksheet's TabId, saves the workbook to a MemoryStream as XLSX, reloads it, captures the TabIds again, and outputs any differences to confirm that TabId values are preserved across serialization.
class TabIdConsistencyCheck
{
    static void Main()
    {
        // Create a new workbook and add some worksheets
        Workbook originalWorkbook = new Workbook(); // create rule
        originalWorkbook.Worksheets.Add("First");
        originalWorkbook.Worksheets.Add("Second");
        originalWorkbook.Worksheets.Add("Third");

        // Capture TabId values before serialization
        Dictionary<string, int> beforeIds = new Dictionary<string, int>();
        foreach (Worksheet ws in originalWorkbook.Worksheets)
        {
            beforeIds[ws.Name] = ws.TabId;
        }

        // Serialize the workbook to a memory stream
        using (MemoryStream ms = new MemoryStream())
        {
            originalWorkbook.Save(ms, SaveFormat.Xlsx); // save rule
            ms.Position = 0; // reset stream position for loading

            // Load the workbook back from the stream
            Workbook loadedWorkbook = new Workbook(ms); // load rule

            // Capture TabId values after deserialization
            Dictionary<string, int> afterIds = new Dictionary<string, int>();
            foreach (Worksheet ws in loadedWorkbook.Worksheets)
            {
                afterIds[ws.Name] = ws.TabId;
            }

            // Compare TabId values
            bool allMatch = true;
            foreach (var kvp in beforeIds)
            {
                string sheetName = kvp.Key;
                int beforeId = kvp.Value;
                int afterId;
                if (afterIds.TryGetValue(sheetName, out afterId))
                {
                    if (beforeId != afterId)
                    {
                        allMatch = false;
                        Console.WriteLine($"TabId mismatch on sheet '{sheetName}': before={beforeId}, after={afterId}");
                    }
                }
                else
                {
                    allMatch = false;
                    Console.WriteLine($"Sheet '{sheetName}' not found after loading.");
                }
            }

            if (allMatch)
            {
                Console.WriteLine("All TabId values are consistent before and after serialization.");
            }
        }
    }
}
