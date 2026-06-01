using System;
using Aspose.Cells;

namespace AsposeCellsColumnWidthDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate a column with long text that needs a wider column
            cells["B1"].PutValue("Short");
            cells["B2"].PutValue("This is a very long piece of text that will not fit in the default column width");
            cells["B3"].PutValue("Another long text entry to demonstrate custom column width");

            // Set a custom width for column B (index 1) in character units
            // Width must be between 0 and 255 characters
            cells.SetColumnWidth(1, 40.0); // 40 characters wide

            // Alternatively, you can set the width via the Column object
            // worksheet.Cells.Columns[1].Width = 40.0;

            // Save the workbook (lifecycle: save)
            workbook.Save("CustomColumnWidthDemo.xlsx");
        }
    }
}