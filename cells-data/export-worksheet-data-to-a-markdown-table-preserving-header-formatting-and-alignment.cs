using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markdown;

namespace AsposeCellsMarkdownExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data – first row will be used as the table header
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["C1"].PutValue("Quantity");

            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.99);
            sheet.Cells["C2"].PutValue(10);

            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.99);
            sheet.Cells["C3"].PutValue(20);

            // Configure markdown save options
            MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
            {
                // Use the first row of the worksheet as the markdown table header
                TableHeaderType = MarkdownTableHeaderType.FirstRow,
                // Enable column alignment by padding with spaces
                AlignColumnPadding = ' ',
                // Optional: set encoding and line separator
                Encoding = Encoding.UTF8,
                LineSeparator = "\n"
            };

            // Save the workbook as a markdown file using the configured options
            workbook.Save("WorksheetExport.md", mdOptions);

            Console.WriteLine("Worksheet exported to markdown file 'WorksheetExport.md' successfully.");
        }
    }
}