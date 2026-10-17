// Title: Insert a UNC network‑share hyperlink into an Excel cell using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a new workbook with Aspose.Cells, adds a hyperlink to cell A1 pointing to a UNC path (e.g., \\Server\Share\File.pdf), sets custom display text, and saves the file. | Show how to use Aspose.Cells Hyperlinks.Add method to embed a network‑share link with double backslashes and then write the workbook to disk.
// Common Searches: aspnet c# how to add a UNC path hyperlink to an Excel worksheet with Aspose.Cells | example of using Aspose.Cells Hyperlinks.Add for network share links | C# Aspose.Cells create hyperlink to \\Server\Share\Document.pdf in cell A1 | save Excel file with UNC hyperlink using Aspose.Cells for .NET
// Tags: add UNC hyperlink Aspose.Cells | Aspose.Cells Hyperlinks.Add usage | Excel cell hyperlink to network share C# | Aspose.Cells workbook save with hyperlink | double backslash UNC path in Aspose.Cells

using System;
using Aspose.Cells;

// The program creates a new workbook, inserts a hyperlink in cell A1 that points to a UNC network share (\\ServerName\SharedFolder\Document.pdf), sets the cell text to "Open Document", and saves the workbook as HyperlinkExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // UNC path to the network share (use double backslashes)
            string uncPath = @"\\ServerName\SharedFolder\Document.pdf";

            // Add a hyperlink to cell A1 (row 0, column 0)
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hyperlink
            sheet.Hyperlinks.Add(0, 0, 1, 1, uncPath);

            // Set the display text in the cell
            sheet.Cells[0, 0].PutValue("Open Document");

            // Save the workbook to a file
            workbook.Save("HyperlinkExample.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
