// Title: C# – Merge Excel workbooks and retain each sheet’s page‑setup using Aspose.Cells
// Description: Shows how to combine multiple Excel files into a single workbook with Aspose.Cells for .NET, then copy the original PageSetup of every source worksheet to its counterpart in the merged file, preserving margins, orientation, scaling, and other print settings.
// Keywords: Aspose.Cells | C# | Workbook.Combine | PageSetup.Copy | merge workbooks | preserve print layout | Excel page setup | copy worksheet settings | combine Excel files | retain margins
// Common Searches: Aspose.Cells copy page setup after combine | merge Excel workbooks keep print settings C# | how to preserve worksheet page layout when merging files Aspose | Workbook.Combine page setup not copied | C# merge multiple Excel files with original page setup
// Developer Intent: Combine several Excel workbooks into one and automatically copy each source worksheet’s page‑setup to the newly added sheet.
// Use Cases: Consolidate quarterly financial reports into a master workbook while keeping each report’s print margins and orientation. | Assemble department‑specific templates into a single distribution file, ensuring every sheet prints correctly without manual adjustments. | Automate creation of a client‑ready workbook that merges diverse source files and retains their original page‑setup to avoid re‑configuring print options.
// AI Prompts: Generate C# code that merges an array of Excel files with Aspose.Cells and copies the PageSetup of each source worksheet to the merged workbook. | Explain how PageSetup.Copy works after Workbook.Combine, including how to map source sheet indexes to destination sheets. | Provide a step‑by‑step guide for preserving print settings when merging multiple workbooks using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace WorkbookMergeWithPageSetup
{
    // Shows how to combine multiple Excel files into a single workbook with Aspose.Cells for .NET, then copy the original PageSetup of every source worksheet to its counterpart in the merged file, preserving margins, orientation, scaling, and other print settings.
    class Program
    {
        static void Main()
        {
            // Paths of source workbooks to be merged
            string[] sourceFiles = { "Source1.xlsx", "Source2.xlsx", "Source3.xlsx" };

            // Destination workbook – starts empty
            Workbook destinationWorkbook = new Workbook();

            // Iterate through each source workbook
            foreach (string filePath in sourceFiles)
            {
                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(filePath);

                // Remember the current number of worksheets in the destination
                int destSheetStartIndex = destinationWorkbook.Worksheets.Count;

                // Combine the source workbook into the destination workbook
                destinationWorkbook.Combine(sourceWorkbook);

                // After combining, copy the page‑setup settings from each source worksheet
                // to its corresponding newly added worksheet in the destination workbook.
                for (int i = 0; i < sourceWorkbook.Worksheets.Count; i++)
                {
                    Worksheet srcSheet = sourceWorkbook.Worksheets[i];
                    Worksheet destSheet = destinationWorkbook.Worksheets[destSheetStartIndex + i];

                    // Copy page‑setup settings using the PageSetup.Copy method.
                    // A new CopyOptions instance is used (default options).
                    destSheet.PageSetup.Copy(srcSheet.PageSetup, new CopyOptions());
                }
            }

            // Save the merged workbook preserving original page layouts
            destinationWorkbook.Save("MergedWithPageSetup.xlsx");
        }
    }
}
