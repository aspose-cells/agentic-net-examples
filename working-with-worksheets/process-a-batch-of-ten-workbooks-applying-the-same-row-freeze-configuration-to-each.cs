// Title: Batch freeze panes in multiple Excel workbooks using Aspose.Cells for .NET
// Description: C# program that loads ten Excel files, applies the same FreezePanes setting (first 3 rows and first 2 columns) to every worksheet in each workbook, and saves the modified files with a "_Frozen" suffix.
// Keywords: Aspose.Cells FreezePanes C# | batch freeze rows columns Excel | process multiple workbooks .NET | freeze panes all worksheets | Aspose.Cells example GitHub | C# Excel automation | freeze first rows columns programmatically
// Common Searches: how to freeze rows and columns in many Excel files with Aspose.Cells | C# batch apply FreezePanes to multiple workbooks | Aspose.Cells loop through worksheets set freeze panes | save Excel files with suffix after modifying with Aspose | example code for freezing panes in all sheets
// Developer Intent: Apply an identical freeze‑pane configuration to every worksheet across a set of Excel workbooks and persist the changes.
// Use Cases: Standardize report layouts by locking header rows and identifier columns in all sheets of each monthly workbook. | Prepare template workbooks for distribution, ensuring key rows/columns stay visible during scrolling. | Automate preprocessing of uploaded Excel files in a web service, adding consistent freeze panes before analysis.
// AI Prompts: Write C# code that reads a list of Excel files, uses Aspose.Cells to set FreezePanes(row, column, frozenRows, frozenColumns) on all worksheets, and saves each file with a custom suffix. | Show how to modify the batch freeze‑panes sample to process any number of files from a directory instead of a hard‑coded array. | Suggest robust error‑handling and logging strategies for loading, updating, and saving multiple workbooks in a loop with Aspose.Cells.

using System;
using Aspose.Cells;

// C# program that loads ten Excel files, applies the same FreezePanes setting (first 3 rows and first 2 columns) to every worksheet in each workbook, and saves the modified files with a "_Frozen" suffix.
class BatchFreezePanes
{
    static void Main()
    {
        // Paths of the ten source workbooks
        string[] sourceFiles = new string[10]
        {
            "Workbook1.xlsx",
            "Workbook2.xlsx",
            "Workbook3.xlsx",
            "Workbook4.xlsx",
            "Workbook5.xlsx",
            "Workbook6.xlsx",
            "Workbook7.xlsx",
            "Workbook8.xlsx",
            "Workbook9.xlsx",
            "Workbook10.xlsx"
        };

        // Freeze configuration: freeze first 3 rows and first 2 columns
        // FreezePanes(row, column, freezedRows, freezedColumns)
        int row = 3;          // zero‑based index where the split occurs
        int column = 2;       // zero‑based index where the split occurs
        int frozenRows = 3;   // number of rows to keep visible at the top
        int frozenColumns = 2; // number of columns to keep visible at the left

        for (int i = 0; i < sourceFiles.Length; i++)
        {
            // Load the workbook from file
            Workbook workbook = new Workbook(sourceFiles[i]);

            // Apply the same freeze setting to every worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.FreezePanes(row, column, frozenRows, frozenColumns);
            }

            // Save the modified workbook (preserving original name with a suffix)
            string outputFile = System.IO.Path.GetFileNameWithoutExtension(sourceFiles[i]) + "_Frozen.xlsx";
            workbook.Save(outputFile);
        }
    }
}
