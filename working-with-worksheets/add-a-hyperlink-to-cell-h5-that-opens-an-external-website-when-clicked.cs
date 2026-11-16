// Title: Add an external website hyperlink to cell H5 in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add hyperlink to cell H5 that opens a web page | How to insert an external URL hyperlink into a specific Excel cell with Aspose.Cells | C# code example for adding a website link to an Excel worksheet using Aspose.Cells | Create a clickable hyperlink in Excel cell H5 using Aspose.Cells for .NET
// Tags: Aspose.Cells add external hyperlink to cell | C# Excel hyperlink insertion Aspose.Cells | hyperlink cell H5 Aspose.Cells .NET | Aspose.Cells worksheet hyperlink API | save workbook with hyperlink Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new workbook, accesses the first worksheet, sets the display text of cell H5 to "Open Website", adds a hyperlink pointing to https://www.example.com to that cell, and saves the workbook as HyperlinkExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the display text for cell H5
        Cell cell = sheet.Cells["H5"];
        cell.PutValue("Open Website");

        // Add a hyperlink to cell H5 (row index 4, column index 7)
        sheet.Hyperlinks.Add(4, 7, 1, 1, "https://www.example.com");

        // Save the workbook
        workbook.Save("HyperlinkExample.xlsx");
    }
}
