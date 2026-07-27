using System;
using Aspose.Cells;

namespace AsposeCellsColumnWidthDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate the column with long text that needs a wider column
            cells["B1"].PutValue("This is a very long piece of text that will not fit in the default column width");
            cells["B2"].PutValue("Another long text entry to demonstrate custom column width handling");
            cells["B3"].PutValue("Yet another example of lengthy content");

            // Display the current width of column B (index 1)
            Console.WriteLine($"Column B width before adjustment: {cells.GetColumnWidth(1)}");

            // Set a custom width for column B (in character units)
            // Width must be between 0 and 255; 30 characters is enough for the sample text
            cells.SetColumnWidth(1, 30);

            // Alternatively, you could use the Column.Width property:
            // worksheet.Cells.Columns[1].Width = 30;

            // Verify the new width
            Console.WriteLine($"Column B width after adjustment: {cells.GetColumnWidth(1)}");

            // Save the workbook (lifecycle save)
            workbook.Save("CustomColumnWidthDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}