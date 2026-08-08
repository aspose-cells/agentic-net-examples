// Title: Add and Verify Cell Watches (B2, C3) in an Excel Workbook with Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, adds watches for cells B2 and C3 using CellWatches.Add, saves the file, reloads it, iterates the CellWatchCollection, and prints whether each watch is present.
// Keywords: Aspose.Cells cell watch C# | CellWatches.Add example | verify cell watch after save | CellWatchCollection iteration | watch window Excel Aspose | load workbook check watches
// Common Searches: how to add cell watches with Aspose.Cells .NET | check if a cell is in the Watch Window after loading workbook | C# Aspose.Cells verify B2 watch | iterate CellWatchCollection Aspose.Cells
// Developer Intent: Add watches for B2 and C3, persist the workbook, reload it, and confirm the watches exist.
// Use Cases: Ensure critical cells are monitored before distributing a workbook. | Automate validation that required watches survive file I/O operations. | Integrate watch‑list checks into batch processing pipelines for data integrity.
// AI Prompts: Generate C# code that adds a list of cell watches and verifies each after loading the workbook with Aspose.Cells. | Write a method returning true if a specified cell name exists in a loaded workbook's CellWatchCollection. | Explain how to remove a cell watch and update the workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsCellWatchDemo
{
    // Creates a new workbook, adds watches for cells B2 and C3 using CellWatches.Add, saves the file, reloads it, iterates the CellWatchCollection, and prints whether each watch is present.
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();                     // uses Workbook() constructor
            Worksheet sheet = workbook.Worksheets[0];               // get the first worksheet

            // Add sample data (optional, just to have some content)
            sheet.Cells["A1"].PutValue("Demo");

            // ---------- Add cell watches ----------
            // Add watches for cells B2 and C3 using the Add(string) method
            int indexB2 = sheet.CellWatches.Add("B2");
            int indexC3 = sheet.CellWatches.Add("C3");

            // Save the workbook to disk (uses the Save(string) method)
            string filePath = "CellWatchDemo.xlsx";
            workbook.Save(filePath);

            // ---------- Load the saved workbook ----------
            Workbook loadedWorkbook = new Workbook(filePath);        // uses Workbook(string) constructor
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

            // ---------- Verify that the specified cells are in the Watch Window ----------
            bool hasB2 = false;
            bool hasC3 = false;

            // Iterate through the CellWatchCollection
            foreach (CellWatch watch in loadedSheet.CellWatches)
            {
                if (watch.CellName.Equals("B2", StringComparison.OrdinalIgnoreCase))
                    hasB2 = true;
                if (watch.CellName.Equals("C3", StringComparison.OrdinalIgnoreCase))
                    hasC3 = true;
            }

            // Output verification results
            Console.WriteLine($"Watch for B2 present: {hasB2}");
            Console.WriteLine($"Watch for C3 present: {hasC3}");
        }
    }
}
