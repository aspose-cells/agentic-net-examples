// Title: How to encapsulate Aspose.Cells FreezePanes logic in a reusable C# method with row and column parameters
// Common Searches: C# Aspose.Cells method to freeze panes at a specific row and column | parameterized FreezePanes function for Excel worksheets using Aspose.Cells | reuse freeze panes code in Aspose.Cells .NET project | apply dynamic freeze panes in Aspose.Cells workbook based on input indices | validate worksheet before calling FreezePanes in Aspose.Cells C#
// Tags: Aspose.Cells FreezePanes C# method | parameterized worksheet freeze panes | dynamic row column freeze panes Aspose | Excel workbook pane freezing validation | reusable freeze panes utility Aspose.Cells

using System;
using Aspose.Cells;

namespace FreezePanesExample
{
    // The example loads an existing workbook, retrieves the first worksheet, and calls a custom FreezePanes method that validates the worksheet and 1‑based row/column arguments, converts them to zero‑based indices, and invokes Worksheet.FreezePanes to freeze the specified rows and columns before saving the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("Input.xlsx");

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Freeze panes at row 5, column 3 (example)
            FreezePanes(sheet, 5, 3);

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("Output.xlsx");
        }

        /// <param name="sheet">The worksheet on which to apply freeze panes.</param>
        /// <param name="row">The row index (1‑based) where the freeze should start.</param>
        /// <param name="column">The column index (1‑based) where the freeze should start.</param>
        static void FreezePanes(Worksheet sheet, int row, int column)
        {
            // Validate parameters
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (row < 1) throw new ArgumentOutOfRangeException(nameof(row), "Row must be >= 1.");
            if (column < 1) throw new ArgumentOutOfRangeException(nameof(column), "Column must be >= 1.");

            // Convert to zero‑based indices required by Aspose.Cells
            int rowIndex = row - 1;
            int columnIndex = column - 1;

            // Apply freeze panes
            sheet.FreezePanes(rowIndex, columnIndex, rowIndex, columnIndex);
        }
    }
}
