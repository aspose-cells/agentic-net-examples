// Title: Conditional Freeze Panes in Aspose.Cells for .NET – Apply Only When Needed
// Description: Shows how to read a worksheet's current frozen pane with GetFreezedPanes and invoke FreezePanes only when the existing row/column freeze differs, preventing redundant calls and preserving scroll position.
// Keywords: Aspose.Cells | .NET | FreezePanes | GetFreezedPanes | conditional freeze | skip redundant freeze | worksheet freeze state | C# Aspose.Cells example | optimize freeze pane | Excel automation
// Common Searches: Aspose.Cells skip FreezePanes if already set | GetFreezedPanes C# example | how to check existing frozen rows Aspose.Cells | conditional FreezePanes .NET | avoid duplicate FreezePanes call | Aspose.Cells freeze pane optimization | compare freeze pane settings Aspose.Cells
// Developer Intent: Detect the current frozen pane configuration and call FreezePanes only when the desired start cell, frozen rows, or frozen columns are different.
// Use Cases: Preserve user scroll position when regenerating a report workbook. | Validate and keep freeze settings in a template before populating data. | Batch‑export multiple worksheets while retaining predefined pane layouts. | Encapsulate the conditional logic in a reusable helper method for projects.
// AI Prompts: Write a C# method that accepts a Worksheet and desired freeze parameters, checks GetFreezedPanes, and applies FreezePanes only if needed. | Explain the values returned by GetFreezedPanes and how to compare them with target freeze settings. | Generate a unit test for the conditional freeze logic using Aspose.Cells and NUnit.

using System;
using Aspose.Cells;

namespace AsposeCellsFreezePaneDemo
{
    // Shows how to read a worksheet's current frozen pane with GetFreezedPanes and invoke FreezePanes only when the existing row/column freeze differs, preventing redundant calls and preserving scroll position.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Desired freeze pane settings
            // Freeze at cell C3 (row index 2, column index 2) with 2 frozen rows and 2 frozen columns
            int desiredRow = 2;          // zero‑based row index where the freeze starts
            int desiredColumn = 2;       // zero‑based column index where the freeze starts
            int desiredFrozenRows = 2;   // number of rows to freeze
            int desiredFrozenColumns = 2;// number of columns to freeze

            // Check current freeze pane state
            int currentRow, currentColumn, currentFrozenRows, currentFrozenColumns;
            bool hasFreeze = worksheet.GetFreezedPanes(out currentRow, out currentColumn,
                                                       out currentFrozenRows, out currentFrozenColumns);

            // Determine whether we need to apply FreezePanes
            bool needToFreeze = true;

            if (hasFreeze)
            {
                // If the existing freeze matches the desired state, skip freezing
                if (currentRow == desiredRow &&
                    currentColumn == desiredColumn &&
                    currentFrozenRows == desiredFrozenRows &&
                    currentFrozenColumns == desiredFrozenColumns)
                {
                    needToFreeze = false;
                }
            }

            // Apply FreezePanes only when necessary
            if (needToFreeze)
            {
                worksheet.FreezePanes(desiredRow, desiredColumn,
                                      desiredFrozenRows, desiredFrozenColumns);
            }

            // Save the workbook (lifecycle rule)
            workbook.Save("FreezePaneResult.xlsx");
        }
    }
}
