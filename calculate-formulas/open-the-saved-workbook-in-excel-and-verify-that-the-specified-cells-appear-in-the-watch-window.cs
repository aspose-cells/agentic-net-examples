// Title: Add cell watches for B2 and E4, save workbook, reload and confirm watches with Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a new workbook, creates a watch for cell B2 and another for cell E4 using Aspose.Cells, saves the file, reloads it, and checks that both watches are present. | Write a C# snippet that iterates through the CellWatchCollection of a loaded worksheet and returns true if a watch for a specific cell name or row/column index is found.
// Common Searches: Aspose.Cells C# register watch for cell B2 and verify after saving | How to determine if a cell watch exists in a loaded workbook with Aspose.Cells | C# example iterating CellWatchCollection to locate watch by row and column | Preserving cell watches when saving and reloading an Excel file using Aspose.Cells
// Tags: add cell watch Aspose.Cells C# | cell watch verification after workbook save | iterate CellWatchCollection Aspose.Cells | preserve cell watches when reloading workbook | watch window handling in Excel via Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, registers watches for cells B2 and E4, saves it as CellWatchDemo.xlsx, reloads the file, iterates through the CellWatchCollection to confirm both watches are retained, and outputs the verification results.
class Program
{
    static void Main()
    {
        // ---------- Create workbook and add cell watches ----------
        Workbook workbook = new Workbook();                         // Workbook()
        Worksheet sheet = workbook.Worksheets[0];                  // access first worksheet

        // Put sample data into a cell (optional, just to have content)
        sheet.Cells["B2"].PutValue("Sample Data");

        // Add watches: one by cell name, another by row/column indices (0‑based)
        int watchIndex1 = sheet.CellWatches.Add("B2");             // Add(string)
        int watchIndex2 = sheet.CellWatches.Add(3, 4);             // Add(int row, int column) -> cell E4

        // Save the workbook to disk
        string fileName = "CellWatchDemo.xlsx";
        workbook.Save(fileName);                                   // Save(string)

        // ---------- Load the saved workbook ----------
        Workbook loadedWorkbook = new Workbook(fileName);           // Workbook(string)
        Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

        // ---------- Verify that the specified cells are in the Watch Window ----------
        bool hasB2 = false;
        bool hasE4 = false;

        foreach (CellWatch watch in loadedSheet.CellWatches)      // iterate over CellWatchCollection
        {
            if (watch.CellName == "B2")
                hasB2 = true;

            // E4 corresponds to row index 3 and column index 4 (0‑based)
            if (watch.Row == 3 && watch.Column == 4)
                hasE4 = true;
        }

        Console.WriteLine($"Watch for B2 present: {hasB2}");
        Console.WriteLine($"Watch for E4 present: {hasE4}");
    }
}
