// Title: Export an Excel worksheet to a Markdown table with header row and aligned columns using Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to write a worksheet to a Markdown (.md) file, using the first row as the table header and aligning columns with space padding. | Show how to configure MarkdownSaveOptions in Aspose.Cells to keep header formatting and column alignment when exporting a workbook to Markdown.
// Common Searches: asp.net export excel worksheet to markdown table with header row using Aspose.Cells | c# markdown table column alignment Aspose.Cells MarkdownSaveOptions example | how to preserve column alignment when saving workbook as .md with Aspose.Cells | export worksheet data to markdown with first row as header Aspose.Cells .NET
// Tags: Aspose.Cells markdown export column padding | C# worksheet to markdown table conversion | MarkdownSaveOptions first row header Aspose | aligned markdown columns from Excel data | export workbook to .md using Aspose.Cells

using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markdown;

namespace AsposeCellsMarkdownExport
{
    // The sample creates a workbook, fills it with product data, sets MarkdownSaveOptions to treat the first row as the table header and pad columns with spaces for alignment, and saves the worksheet as a UTF-8 Markdown file named ExportedTable.md.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate worksheet with sample data (including a header row)
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Price");
            worksheet.Cells["C1"].PutValue("Quantity");

            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(1.99);
            worksheet.Cells["C2"].PutValue(10);

            worksheet.Cells["A3"].PutValue("Banana");
            worksheet.Cells["B3"].PutValue(0.99);
            worksheet.Cells["C3"].PutValue(20);

            // Configure markdown save options:
            // - Use the first row as the table header
            // - Enable column alignment by padding with spaces
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
            {
                TableHeaderType = MarkdownTableHeaderType.FirstRow,
                AlignColumnPadding = ' ',
                Encoding = Encoding.UTF8,
                LineSeparator = "\n"
            };

            // Save the workbook as a markdown file preserving header formatting and alignment
            workbook.Save("ExportedTable.md", saveOptions);

            Console.WriteLine("Worksheet exported to markdown successfully.");
        }
    }
}
