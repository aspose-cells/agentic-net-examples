// Title: Add a hyperlink to an external PDF in a cell and freeze the header row with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add hyperlink to PDF file in Excel cell | How to freeze the first row after inserting a hyperlink using Aspose.Cells | C# code to create a hyperlink to a local document and keep header row static in Excel | Aspose.Cells example for adding external document link and freezing panes | Freeze panes on top row while linking to external PDF in Aspose.Cells workbook
// Tags: cell hyperlink insertion Aspose.Cells C# | freeze first row panes Aspose.Cells | link external PDF file Aspose.Cells | worksheet freeze panes after adding hyperlink | Aspose.Cells hyperlink freeze example

using Aspose.Cells;
using System;

// The sample creates a new workbook, writes "Open Document" to cell A1, adds a hyperlink to a local PDF file on that cell, freezes the first worksheet row, and saves the workbook as HyperlinkAndFreeze.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put display text into cell A1
            sheet.Cells["A1"].PutValue("Open Document");

            // Add a hyperlink to the external document (cell A1)
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hyperlink address
            sheet.Hyperlinks.Add(0, 0, 1, 1, @"C:\Docs\example.pdf");

            // Freeze the first row (rows above row index 1)
            // Parameters: row, column, totalRows, totalColumns
            sheet.FreezePanes(1, 0, 0, 0);

            // Save the workbook
            workbook.Save("HyperlinkAndFreeze.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
