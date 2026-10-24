// Title: Verify that each worksheet keeps a unique Index (SheetId) after using AddCopy in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# check for duplicate worksheet indexes after AddCopy | how to confirm worksheet IDs are unique after copying a sheet in .NET | detect duplicate sheet indices in a workbook using Aspose.Cells API | C# code to validate unique worksheet Index values after copying worksheets | ensure worksheet Index remains distinct when duplicating sheets with Aspose.Cells
// Tags: AddCopy worksheet index uniqueness Aspose.Cells | C# verify unique worksheet IDs | duplicate sheet index detection Aspose.Cells | HashSet validation of worksheet indices | Aspose.Cells workbook integrity check

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example creates a workbook, adds two worksheets, copies one with AddCopy, and then uses a HashSet to iterate through all worksheets, confirming that each worksheet's Index (used as SheetId) is unique and reporting any duplicates.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add first worksheet and set a name
            int sheet1Index = workbook.Worksheets.Add();
            Worksheet sheet1 = workbook.Worksheets[sheet1Index];
            sheet1.Name = "Sheet1";

            // Add second worksheet and set a name
            int sheet2Index = workbook.Worksheets.Add();
            Worksheet sheet2 = workbook.Worksheets[sheet2Index];
            sheet2.Name = "Sheet2";

            // Copy the first worksheet to a new worksheet
            int copiedIndex = workbook.Worksheets.AddCopy(sheet1.Index);
            Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
            copiedSheet.Name = "Sheet1_Copy";

            // Validate that each worksheet retains a unique Index (acts as SheetId)
            HashSet<int> sheetIndices = new HashSet<int>();
            foreach (Worksheet ws in workbook.Worksheets)
            {
                int idx = ws.Index;
                // If the Index already exists in the set, it's a duplicate
                if (!sheetIndices.Add(idx))
                {
                    Console.WriteLine($"Duplicate worksheet Index detected: {idx} on worksheet '{ws.Name}'.");
                    return;
                }
            }

            Console.WriteLine("All worksheets have unique indices.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
