// Title: Apply Saved Watch Window (CellWatch) to Another Workbook using Aspose.Cells C#
// Description: Loads a source workbook that contains a saved CellWatch (Watch Window) configuration, ensures the target workbook has matching worksheets, copies each CellWatch entry (cell name, row, column) to the corresponding worksheet, and saves the target file with the applied watch settings. Ideal for automating monitoring of critical cells across Excel reports.
// Keywords: Aspose.Cells | C# | CellWatch | Watch Window | copy watch configuration | apply saved watch settings | Excel monitoring automation | transfer CellWatch between workbooks | load watch window configuration | Aspose.Cells tutorial
// Common Searches: How to copy CellWatch from one Excel file to another with Aspose.Cells | Apply saved Watch Window settings to a target workbook C# | Transfer watch window entries between workbooks using Aspose.Cells | Copy watch window cells programmatically Aspose.Cells .NET | Load and apply CellWatch configuration in C#
// Developer Intent: Copy a saved CellWatch (Watch Window) configuration from a source workbook and apply it to a target workbook using Aspose.Cells for .NET.
// Use Cases: Reuse a template's watch list in newly generated reports to monitor key cells automatically. | Synchronize watch windows across multiple workbooks in a batch reporting pipeline. | Consolidate watch configurations from several source files into a master workbook for centralized monitoring.
// AI Prompts: Generate C# code with Aspose.Cells that copies all CellWatch objects from a source workbook to a target workbook, adding missing worksheets as needed. | Explain best practices for transferring Watch Window entries when source and target workbooks have different worksheet counts. | Create a reusable method that accepts source and target file paths, copies CellWatch settings, and returns the updated target workbook.

using System;
using System.IO;
using Aspose.Cells;

// Loads a source workbook that contains a saved CellWatch (Watch Window) configuration, ensures the target workbook has matching worksheets, copies each CellWatch entry (cell name, row, column) to the corresponding worksheet, and saves the target file with the applied watch settings. Ideal for automating monitoring of critical cells across Excel reports.
class ApplyWatchWindow
{
    static void Main()
    {
        // Paths for source (contains saved watch configuration) and target workbooks
        string sourcePath = "SourceWithWatch.xlsx";
        string targetPath = "TargetWorkbook.xlsx";

        try
        {
            // Load source workbook if it exists; otherwise create an empty workbook
            Workbook sourceWb = File.Exists(sourcePath) ? new Workbook(sourcePath) : new Workbook();

            // Load target workbook if it exists; otherwise create a new workbook
            Workbook targetWb = File.Exists(targetPath) ? new Workbook(targetPath) : new Workbook();

            // Ensure the target workbook has at least the same number of worksheets as the source
            while (targetWb.Worksheets.Count < sourceWb.Worksheets.Count)
            {
                targetWb.Worksheets.Add();
            }

            // Copy CellWatch settings from each source worksheet to the corresponding target worksheet
            for (int i = 0; i < sourceWb.Worksheets.Count; i++)
            {
                Worksheet srcSheet = sourceWb.Worksheets[i];
                Worksheet tgtSheet = targetWb.Worksheets[i];

                foreach (CellWatch srcWatch in srcSheet.CellWatches)
                {
                    // Add the same cell name to the target sheet's watch collection
                    int watchIndex = tgtSheet.CellWatches.Add(srcWatch.CellName);

                    // Copy additional properties (row, column, name) to the new watch
                    CellWatch tgtWatch = tgtSheet.CellWatches[watchIndex];
                    tgtWatch.Row = srcWatch.Row;
                    tgtWatch.Column = srcWatch.Column;
                    tgtWatch.CellName = srcWatch.CellName;
                }
            }

            // Save the target workbook with the applied Watch Window configuration
            targetWb.Save("TargetWorkbook_WithWatch.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
