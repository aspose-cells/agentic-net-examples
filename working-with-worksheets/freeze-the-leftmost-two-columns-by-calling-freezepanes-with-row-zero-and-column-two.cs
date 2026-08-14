// Title: Freeze the First Two Columns in Excel with Aspose.Cells for .NET
// Description: Demonstrates how to use Aspose.Cells' Worksheet.FreezePanes method (row 0, column 2) to lock the leftmost two columns (A‑B) while keeping all rows unfrozen, then saves the workbook as FreezeLeftmostTwoColumns.xlsx.
// Keywords: Aspose.Cells | C# | FreezePanes | freeze columns Excel | leftmost columns | programmatic freeze | .NET Excel API | worksheet FreezePanes example
// Common Searches: Aspose.Cells freeze first two columns C# | How to lock leftmost columns in Excel using Aspose.Cells | Worksheet.FreezePanes row 0 column 2 example | Freeze columns without freezing rows Aspose.Cells .NET | C# code to freeze columns in an Excel workbook
// Developer Intent: Lock columns A and B in a worksheet while leaving all rows scrollable.
// Use Cases: Financial reports where identifier columns must stay visible during horizontal scrolling. | Data‑entry templates that keep index columns fixed to avoid accidental edits. | Large datasets where key columns need to remain in view while reviewing rows.
// AI Prompts: Show how to freeze the first three columns of a worksheet using Aspose.Cells for .NET. | Explain each parameter of Worksheet.FreezePanes and how to freeze rows independently of columns. | Provide C# code that validates a worksheet object before calling FreezePanes.

using System;
using Aspose.Cells;

namespace FreezeLeftmostColumnsDemo
{
    // Demonstrates how to use Aspose.Cells' Worksheet.FreezePanes method (row 0, column 2) to lock the leftmost two columns (A‑B) while keeping all rows unfrozen, then saves the workbook as FreezeLeftmostTwoColumns.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Freeze the leftmost two columns.
            // Parameters: row = 0 (no frozen rows), column = 2 (freeze at column C),
            // freezedRows = 0, freezedColumns = 2 (freeze two columns on the left)
            worksheet.FreezePanes(0, 2, 0, 2);

            // Save the workbook to a file
            workbook.Save("FreezeLeftmostTwoColumns.xlsx");
        }
    }
}
