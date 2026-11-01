// Title: Freeze the top row and left column in an Aspose.Cells worksheet and export to HTML with persistent headers (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to apply FreezePanes on the first row and column, then saves the workbook as an HTML file with those panes rendered as static headers. | Show a complete example of freezing header rows/columns in a worksheet and exporting it to HTML using Aspose.Cells for .NET, ensuring the HTML view keeps the headers fixed.
// Common Searches: Aspose.Cells C# export to HTML while preserving frozen panes as header rows | How to keep frozen rows and columns visible after saving Excel as HTML with Aspose.Cells | C# example of FreezePanes followed by HTML output using Aspose.Cells library | Saving an Excel worksheet as HTML with fixed header using Aspose.Cells .NET
// Tags: header row freezing Aspose.Cells C# | HTML export with static headers Aspose.Cells | Aspose.Cells FreezePanes API usage | SaveFormat.Html with frozen headers Aspose.Cells | preserve worksheet frozen panes in HTML Aspose.Cells

using Aspose.Cells;
using System;

// The program creates a workbook, fills it with sample data, freezes the first row and column using FreezePanes, and saves the worksheet as an HTML file where the frozen panes appear as fixed header rows and columns.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Fill the worksheet with sample data (optional)
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Freeze the first row and first column (headers)
            // FreezePanes(row, column, totalRows, totalColumns)
            int totalRows = sheet.Cells.MaxDataRow + 1;
            int totalColumns = sheet.Cells.MaxDataColumn + 1;
            sheet.FreezePanes(1, 1, totalRows, totalColumns);

            // Export the worksheet to HTML; frozen panes appear as fixed headers
            workbook.Save("FrozenHeaders.html", SaveFormat.Html);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
