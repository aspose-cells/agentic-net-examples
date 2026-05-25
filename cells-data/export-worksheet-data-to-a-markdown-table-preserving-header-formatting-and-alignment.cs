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
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data (first row will be used as header)
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Price");
            worksheet.Cells["C1"].PutValue("Quantity");

            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(1.99);
            worksheet.Cells["C2"].PutValue(10);

            worksheet.Cells["A3"].PutValue("Banana");
            worksheet.Cells["B3"].PutValue(0.99);
            worksheet.Cells["C3"].PutValue(20);

            // Configure markdown save options
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
            {
                // Use the first row as the table header to preserve header formatting
                TableHeaderType = MarkdownTableHeaderType.FirstRow,
                // Enable column alignment by padding with spaces
                AlignColumnPadding = ' ',
                // Optional: set encoding and line separator
                Encoding = Encoding.UTF8,
                LineSeparator = "\n"
            };

            // Save the workbook as a markdown file
            string outputPath = "WorksheetExport.md";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Worksheet exported to markdown file: {outputPath}");
        }
    }
}