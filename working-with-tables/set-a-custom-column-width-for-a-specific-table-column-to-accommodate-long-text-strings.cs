// Title: C# – Set Custom Column Width for a Specific Column with Aspose.Cells
// Description: Shows how to create a workbook, insert a long string into cell B1, and assign a fixed width of 40 characters to column B using Cells.SetColumnWidth (or the Columns collection) before saving the file as CustomColumnWidth.xlsx.
// Keywords: Aspose.Cells | C# | set column width | custom column width | column width characters | Excel column width programmatically | Cells.SetColumnWidth | worksheet column width | adjust column width for long text | Aspose.Cells .NET
// Common Searches: Aspose.Cells set column width C# | how to change column width in Excel using Aspose.Cells | set column width in characters with Aspose.Cells | increase column width for long text Aspose.Cells .NET | programmatically adjust column width in a workbook
// Developer Intent: Define a fixed width for a chosen worksheet column so that lengthy text is fully visible without truncation.
// Use Cases: Set column B to 40 characters to accommodate description fields in generated reports. | Apply column width via Cells.SetColumnWidth or worksheet.Cells.Columns[index].Width before exporting the workbook. | Create Excel files with predefined column sizes for tables that contain long strings.
// AI Prompts: Write C# code that automatically sets a column's width based on the longest string in that column using Aspose.Cells. | Provide an example that loops through multiple columns and assigns appropriate widths before saving the workbook. | Explain the difference between character units and points when setting column width with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsColumnWidthExample
{
    // Shows how to create a workbook, insert a long string into cell B1, and assign a fixed width of 40 characters to column B using Cells.SetColumnWidth (or the Columns collection) before saving the file as CustomColumnWidth.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate a cell with a long text string in column B (index 1)
            cells["B1"].PutValue("This is a very long piece of text that needs a wider column to be fully visible.");

            // Set a custom width for column B (index 1) in character units.
            // Width must be between 0 and 255. Adjust as needed for the text length.
            cells.SetColumnWidth(1, 40.0); // 40 characters wide

            // Optionally, you can also set the width using the Columns collection:
            // worksheet.Cells.Columns[1].Width = 40.0;

            // Save the workbook (lifecycle rule: save)
            workbook.Save("CustomColumnWidth.xlsx", SaveFormat.Xlsx);
        }
    }
}
