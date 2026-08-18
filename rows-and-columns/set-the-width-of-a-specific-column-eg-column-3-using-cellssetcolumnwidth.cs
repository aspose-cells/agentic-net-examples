// Title: Set column 3 width to 25.5 characters with Aspose.Cells for .NET (C#)
// Description: Creates a new Workbook, accesses the first Worksheet, and uses Cells.SetColumnWidth(2, 25.5) to set the third column (zero‑based index 2) to 25.5 character units before saving as ColumnWidthDemo.xlsx.
// Keywords: Aspose.Cells C# column width | SetColumnWidth example | adjust Excel column width .NET | column width by index Aspose | Excel column size characters
// Common Searches: Aspose.Cells set column width C# | how to change column 3 width Aspose.Cells | SetColumnWidth zero based index .NET | adjust column width before saving Excel with Aspose
// Developer Intent: Define the width of a specific worksheet column using Aspose.Cells.
// Use Cases: Improve readability of generated reports by fixing column widths. | Prepare consistent layouts for printed Excel documents. | Standardize column dimensions across multiple worksheets in automated exports.
// AI Prompts: Write C# code that sets custom widths for columns 1‑5 using Aspose.Cells and saves the workbook. | Explain the measurement unit used by Cells.SetColumnWidth and how to convert pixels to character units. | Create a reusable method that takes a worksheet, column index, and width, applies SetColumnWidth, and includes error handling.

using System;
using Aspose.Cells;

namespace AsposeCellsColumnWidthExample
{
    // Creates a new Workbook, accesses the first Worksheet, and uses Cells.SetColumnWidth(2, 25.5) to set the third column (zero‑based index 2) to 25.5 character units before saving as ColumnWidthDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the Cells collection
            Cells cells = worksheet.Cells;

            // Set the width of column 3 (zero‑based index 2) to 25.5 characters
            cells.SetColumnWidth(2, 25.5);

            // Save the workbook (lifecycle: save)
            workbook.Save("ColumnWidthDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
