// Title: Check that Worksheet TabId values remain unchanged after saving and reloading an XLSX workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that records each worksheet's TabId, saves the workbook to a MemoryStream in XLSX format, reloads it, and verifies the TabId values are identical. | Show how to compare two List<int> of Worksheet.TabId before and after a workbook round‑trip with Aspose.Cells.
// Common Searches: how to ensure worksheet TabId stays the same after Aspose.Cells workbook save and load in C# | Aspose.Cells .NET compare TabId values before and after workbook serialization | preserve worksheet tab order when saving Excel file with Aspose.Cells | C# verify worksheet TabId consistency after round‑trip XLSX using Aspose.Cells
// Tags: worksheet TabId verification after XLSX roundtrip | Aspose.Cells workbook serialization to MemoryStream | compare Worksheet.TabId collections in C# | preserve Excel tab identifiers with Aspose.Cells | validate worksheet order after workbook reload

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The program creates a workbook with multiple worksheets, captures each worksheet's TabId, saves the workbook to a MemoryStream as XLSX, reloads it, captures the TabId values again, and compares the two sets to confirm that TabId values remain consistent across the serialization round‑trip.
class Program
{
    static void Main()
    {
        try
        {
            // ---------- Create workbook ----------
            var workbook = new Workbook();

            // Rename the default worksheet to avoid name conflict
            workbook.Worksheets[0].Name = "Sheet0";

            // Add additional worksheets with unique names
            workbook.Worksheets.Add("Sheet1");
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Capture TabId values before serialization
            var originalTabIds = new List<int>();
            foreach (Worksheet ws in workbook.Worksheets)
            {
                originalTabIds.Add(ws.TabId);
            }

            // ---------- Serialize workbook ----------
            using (var ms = new MemoryStream())
            {
                // Save to memory stream (serialization)
                workbook.Save(ms, SaveFormat.Xlsx);
                ms.Position = 0; // reset stream for reading

                // ---------- Load workbook ----------
                var loadedWorkbook = new Workbook(ms);

                // Capture TabId values after deserialization
                var loadedTabIds = new List<int>();
                foreach (Worksheet ws in loadedWorkbook.Worksheets)
                {
                    loadedTabIds.Add(ws.TabId);
                }

                // ---------- Compare TabId values ----------
                bool isConsistent = originalTabIds.Count == loadedTabIds.Count;
                if (isConsistent)
                {
                    for (int i = 0; i < originalTabIds.Count; i++)
                    {
                        if (originalTabIds[i] != loadedTabIds[i])
                        {
                            isConsistent = false;
                            break;
                        }
                    }
                }

                Console.WriteLine("TabId consistency after serialization: " + isConsistent);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
