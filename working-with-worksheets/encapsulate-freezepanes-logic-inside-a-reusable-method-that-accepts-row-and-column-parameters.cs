// Title: Reusable FreezePanes Extension Method for Aspose.Cells (C#)
// Description: Provides a static Worksheet extension method FreezeAt that wraps Aspose.Cells' FreezePanes call. The method accepts zero‑based row and column indices, uses them as both the freeze line and the count of rows/columns to lock, and is demonstrated by freezing the first three rows and two columns before saving the workbook.
// Keywords: Aspose.Cells | C# extension method | FreezePanes | worksheet freeze panes | reusable utility | Excel header freeze | zero‑based indices | FreezeAt method | Excel export helper | Aspose.Cells API
// Common Searches: Aspose.Cells how to freeze panes with a helper method | C# extension to lock rows and columns in Excel | Reusable FreezePanes code sample Aspose | Freeze first rows and columns using Aspose.Cells C# | Worksheet FreezeAt utility example
// Developer Intent: Create a single, reusable method that freezes rows and columns based on supplied indices, eliminating repetitive FreezePanes calls.
// Use Cases: Standardize header row and column freezing across all generated worksheets. | Simplify Excel report generation by calling sheet.FreezeAt(row, col) instead of raw FreezePanes parameters. | Encapsulate pane‑freezing logic in a shared library for multiple .NET projects using Aspose.Cells.
// AI Prompts: Generate a C# extension method named FreezeAt for Aspose.Cells Worksheet that takes zero‑based row and column indices and internally calls FreezePanes with matching frozen rows and columns, including XML documentation. | Show example code that uses FreezeAt to lock the first 5 rows and 3 columns of a worksheet and then saves the file as Report.xlsx. | Explain how to integrate the FreezeAt helper into an existing Aspose.Cells solution and apply it automatically to every worksheet in a workbook.

using System;
using Aspose.Cells;

namespace FreezePanesUtility
{
    // Provides a static Worksheet extension method FreezeAt that wraps Aspose.Cells' FreezePanes call. The method accepts zero‑based row and column indices, uses them as both the freeze line and the count of rows/columns to lock, and is demonstrated by freezing the first three rows and two columns before saving the workbook.
    public static class PaneHelper
    {
        /// <param name="worksheet">Target worksheet.</param>
        /// <param name="row">Zero‑based row index where the freeze line starts.</param>
        /// <param name="column">Zero‑based column index where the freeze line starts.</param>
        public static void FreezeAt(this Worksheet worksheet, int row, int column)
        {
            // FreezePanes(row, column, freezedRows, freezedColumns)
            // Here we freeze the same number of rows and columns as the position,
            // which is the most common scenario.
            worksheet.FreezePanes(row, column, row, column);
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook (using the standard creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Example: freeze panes at row 3, column 2 (zero‑based indices)
            // This will freeze the first three rows and first two columns.
            sheet.FreezeAt(3, 2);

            // Save the workbook (using the standard save rule)
            workbook.Save("FreezePanesResult.xlsx");
        }
    }
}
