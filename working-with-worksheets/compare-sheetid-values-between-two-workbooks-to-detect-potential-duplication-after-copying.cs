// Title: Detect duplicate worksheet indexes between two Excel workbooks using Aspose.Cells for .NET
// Common Searches: how to find duplicate worksheet indexes in two Excel files with Aspose.Cells C# | compare worksheet Index property across multiple workbooks using Aspose.Cells | C# code to check for overlapping sheet IDs after copying workbooks with Aspose.Cells | Aspose.Cells detect same worksheet index in different workbooks | identify duplicate sheet indexes when merging Excel workbooks in .NET
// Tags: worksheet index comparison Aspose.Cells | duplicate sheet detection C# | hashset worksheet index validation Aspose.Cells | compare sheet indexes Excel .NET | Aspose.Cells duplicate worksheet identification

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace SheetIdComparison
{
    // The example loads two Excel workbooks with Aspose.Cells, stores each worksheet's Index from the first workbook in a HashSet, then iterates the second workbook to flag any worksheets that share the same Index, indicating possible duplicate sheet IDs.
    class SheetIdComparer
    {
        static void Main()
        {
            try
            {
                string workbookPath1 = "Workbook1.xlsx";
                string workbookPath2 = "Workbook2.xlsx";

                // Verify that the input files exist
                if (!File.Exists(workbookPath1))
                {
                    Console.WriteLine($"File not found: {workbookPath1}");
                    return;
                }
                if (!File.Exists(workbookPath2))
                {
                    Console.WriteLine($"File not found: {workbookPath2}");
                    return;
                }

                // Load the workbooks
                Workbook wb1 = new Workbook(workbookPath1);
                Workbook wb2 = new Workbook(workbookPath2);

                // Collect worksheet indexes from the first workbook
                HashSet<int> sheetIndexes = new HashSet<int>();
                foreach (Worksheet ws in wb1.Worksheets)
                {
                    sheetIndexes.Add(ws.Index); // Use Index as a unique identifier within a workbook
                }

                // Compare worksheet indexes in the second workbook against the first workbook
                bool duplicationFound = false;
                foreach (Worksheet ws in wb2.Worksheets)
                {
                    if (sheetIndexes.Contains(ws.Index))
                    {
                        Console.WriteLine($"Potential duplication detected: Sheet '{ws.Name}' in Workbook2 has Index {ws.Index}, which also exists in Workbook1.");
                        duplicationFound = true;
                    }
                }

                if (!duplicationFound)
                {
                    Console.WriteLine("No duplicate worksheet indexes were found between the two workbooks.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
