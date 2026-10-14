// Title: Define custom page margins and freeze the top row in an Excel worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to set left, right, top, bottom, header, and footer margins in inches and freeze the first worksheet row. | Create a .NET example that configures page margins, applies FreezePanes to keep the header row visible, and saves the workbook as an XLSX file.
// Common Searches: Aspose.Cells C# set page margins in inches and freeze first row | How to freeze top row after defining custom margins with Aspose.Cells .NET | C# example for configuring header and footer margins and FreezePanes in Aspose.Cells | Saving workbook with custom printable area and frozen header using Aspose.Cells | Aspose.Cells page setup margins and freeze panes code sample
// Tags: page margins Aspose.Cells C# | freeze panes top row Aspose.Cells | header footer margin configuration Aspose.Cells | save workbook as xlsx Aspose.Cells | custom printable area Aspose.Cells | worksheet page setup C# Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsMarginAndFreezeExample
{
    // The program creates a new workbook, sets custom left, right, top, bottom, header, and footer margins in inches, freezes the first worksheet row with FreezePanes, and saves the file as MarginAndFreezeDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------
                // Define custom page margins (in inches)
                // -------------------------
                // Set left, right, top, and bottom margins
                sheet.PageSetup.LeftMargin = 0.5;    // 0.5 inch left margin
                sheet.PageSetup.RightMargin = 0.5;   // 0.5 inch right margin
                sheet.PageSetup.TopMargin = 0.75;    // 0.75 inch top margin
                sheet.PageSetup.BottomMargin = 0.75; // 0.75 inch bottom margin

                // Optionally set header and footer margins
                sheet.PageSetup.HeaderMargin = 0.3;
                sheet.PageSetup.FooterMargin = 0.3;

                // -------------------------
                // Freeze the top row so it stays visible and within the printable area
                // -------------------------
                // FreezePanes(row, column, totalRows, totalColumns)
                // Freeze the first row (totalRows = 1) and no columns
                sheet.FreezePanes(1, 0, 1, 0);

                // Save the workbook to a file
                workbook.Save("MarginAndFreezeDemo.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
