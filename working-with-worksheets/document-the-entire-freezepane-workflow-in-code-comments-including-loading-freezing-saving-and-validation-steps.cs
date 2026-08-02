// Title: Aspose.Cells .NET Freeze Panes Example – Load, Freeze, Validate, Save as XLSB
// Description: C# sample that creates or loads a workbook, freezes the first three rows and columns at cell C3, reads back the freeze parameters with GetFreezedPanes, checks PaneState, and saves the file as XLSB using XlsbSaveOptions with merged‑area validation.
// Keywords: Aspose.Cells | C# | .NET | freeze panes | Worksheet.FreezePanes | GetFreezedPanes | PaneState | XlsbSaveOptions | ValidateMergedAreas | XLSB | Excel automation
// Common Searches: Aspose.Cells freeze first rows columns C# | GetFreezedPanes example Aspose.Cells | save XLSB with merged area validation Aspose.Cells | PaneState enumeration Aspose.Cells | C# code to freeze panes in Excel | Aspose.Cells freeze pane workflow
// Developer Intent: Demonstrate how to programmatically freeze rows/columns, verify the settings, and persist the workbook with validation using Aspose.Cells for .NET.
// Use Cases: Generate reports where header rows and columns stay visible while scrolling. | Automate Excel creation on a server that requires frozen panes for better navigation. | Validate merged cells before saving to XLSB to avoid file corruption. | Integrate freeze‑pane logic into a document‑processing pipeline that outputs XLSB files.
// AI Prompts: Write C# code with Aspose.Cells to freeze the first three rows and columns at cell C3 and print the freeze details using GetFreezedPanes. | Show how to save a workbook as .xlsb with ValidateMergedAreas enabled and explain why this setting matters. | Explain the possible values of PaneState after applying FreezePanes and how to interpret each state. | Create a script that loads an existing workbook, applies freeze panes, checks GetFreezedPanes output, and logs the results.

using System;
using Aspose.Cells;

namespace FreezePaneWorkflowDemo
{
    // C# sample that creates or loads a workbook, freezes the first three rows and columns at cell C3, reads back the freeze parameters with GetFreezedPanes, checks PaneState, and saves the file as XLSB using XlsbSaveOptions with merged‑area validation.
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Load or create a workbook
            // ------------------------------------------------------------
            // Here we create a new workbook. In a real scenario you could
            // load an existing file using: new Workbook("input.xlsx");
            Workbook workbook = new Workbook();

            // ------------------------------------------------------------
            // 2. Access the target worksheet
            // ------------------------------------------------------------
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "DataSheet";

            // ------------------------------------------------------------
            // 3. Populate some sample data (optional, just for illustration)
            // ------------------------------------------------------------
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // ------------------------------------------------------------
            // 4. Freeze panes
            // ------------------------------------------------------------
            // Freeze the first three rows and first three columns.
            // The freeze position is cell "C3" (row index 2, column index 2).
            // The last two parameters specify how many rows/columns are frozen.
            worksheet.FreezePanes("C3", 3, 3);

            // ------------------------------------------------------------
            // 5. Validate that the panes are frozen
            // ------------------------------------------------------------
            // GetFreezedPanes returns a boolean indicating whether the
            // worksheet has frozen panes and outputs the freeze details.
            int freezeRow, freezeColumn, frozenRows, frozenColumns;
            bool hasFreeze = worksheet.GetFreezedPanes(out freezeRow, out freezeColumn, out frozenRows, out frozenColumns);

            Console.WriteLine($"Has frozen panes: {hasFreeze}");
            if (hasFreeze)
            {
                Console.WriteLine($"Freeze position - Row: {freezeRow}, Column: {freezeColumn}");
                Console.WriteLine($"Frozen rows: {frozenRows}, Frozen columns: {frozenColumns}");
            }

            // ------------------------------------------------------------
            // 6. Check the overall pane state (optional)
            // ------------------------------------------------------------
            // PaneState provides a higher‑level enumeration of the pane status.
            PaneStateType paneState = worksheet.PaneState;
            Console.WriteLine($"Pane state: {paneState}");

            // ------------------------------------------------------------
            // 7. Save the workbook with validation of merged areas
            // ------------------------------------------------------------
            // Although this example does not merge cells, setting
            // ValidateMergedAreas demonstrates how to enable validation
            // during the save operation.
            XlsbSaveOptions saveOptions = new XlsbSaveOptions
            {
                ValidateMergedAreas = true,
                // Preserve merged areas (default is true, set explicitly for clarity)
                MergeAreas = true
            };

            // Save the workbook to a file.
            workbook.Save("FreezePaneWorkflowDemo.xlsb", saveOptions);

            Console.WriteLine("Workbook saved successfully.");
        }
    }
}
