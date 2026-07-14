using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (using the provided create rule)
        Workbook workbook = new Workbook(); // placeholder for create rule

        // Add sample worksheets
        workbook.Worksheets.Add("SheetA");
        workbook.Worksheets.Add("SheetB");
        workbook.Worksheets.Add("SheetC");

        // Name of the sheet after which the new sheet should be placed
        string targetSheetName = "SheetB";

        // Locate the index of the target sheet
        int targetIndex = -1;
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            if (workbook.Worksheets[i].Name.Equals(targetSheetName, StringComparison.OrdinalIgnoreCase))
            {
                targetIndex = i;
                break;
            }
        }

        if (targetIndex >= 0)
        {
            // Insert a new worksheet immediately after the target sheet
            int insertIndex = targetIndex + 1;
            Worksheet newSheet = workbook.Worksheets.Insert(insertIndex, SheetType.Worksheet, "InsertedAfterTarget");
            // Example content
            newSheet.Cells["A1"].PutValue($"Inserted after {targetSheetName}");
        }

        // Save the workbook (using the provided save rule)
        workbook.Save("SequencedWorkbook.xlsx"); // placeholder for save rule
    }
}

// Author: Example demonstrating how to place a worksheet immediately after a specified sheet name using Aspose.Cells.