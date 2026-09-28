// Title: Generate HTML from an Excel workbook while controlling hidden rows using Aspose.Cells for .NET
// AI Prompts: Create an HTML file from a workbook that excludes any hidden rows using the default HtmlSaveOptions. | Export a workbook to HTML and include hidden rows by setting ExportHiddenWorksheet to true in HtmlSaveOptions. | Programmatically hide specific rows in a worksheet before saving it as HTML with Aspose.Cells.
// Common Searches: Aspose.Cells C# export worksheet to HTML without hidden rows | How to include hidden rows when saving Excel as HTML using Aspose.Cells | HtmlSaveOptions ExportHiddenWorksheet true example C# | Hide rows in a worksheet then export to HTML with Aspose.Cells
// Tags: export hidden rows html Aspose.Cells | HtmlSaveOptions ExportHiddenWorksheet | hide rows before html export C# | Aspose.Cells generate html without hidden rows | include hidden rows in html output Aspose

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, fills cells A1:J10, hides rows 3‑5, saves to HTML with default options (hidden rows omitted), then saves again with ExportHiddenWorksheet set to true so hidden rows appear in the output.
class ExportHiddenRowsExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill some data in rows 0-9 (A1:J10)
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Hide rows 3, 4 and 5 (zero‑based index)
            sheet.Cells.Rows[2].IsHidden = true; // Row 3
            sheet.Cells.Rows[3].IsHidden = true; // Row 4
            sheet.Cells.Rows[4].IsHidden = true; // Row 5

            // Export to HTML with default options (hidden rows are omitted)
            HtmlSaveOptions defaultOptions = new HtmlSaveOptions();
            workbook.Save("Workbook_HiddenRows_Omitted.html", defaultOptions);

            // Export to HTML with ExportHiddenWorksheet = true (hidden rows are included)
            HtmlSaveOptions includeHiddenOptions = new HtmlSaveOptions
            {
                ExportHiddenWorksheet = true // Show hidden rows in the output
            };
            workbook.Save("Workbook_HiddenRows_Included.html", includeHiddenOptions);

            Console.WriteLine("HTML files generated successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
