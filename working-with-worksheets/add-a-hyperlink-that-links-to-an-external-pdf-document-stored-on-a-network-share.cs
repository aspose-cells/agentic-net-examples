// Title: Create an Excel workbook with a UNC‑path hyperlink to a PDF file using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that adds a network‑share hyperlink pointing to a PDF file into cell B2 of a new worksheet using Aspose.Cells. | Write a C# program that creates an Excel file, inserts the text "Open PDF" in B2, attaches a hyperlink to \\ServerName\SharedFolder\Document.pdf, saves the workbook, and prints the absolute output path. | Produce an Aspose.Cells snippet that adds a UNC address hyperlink, saves the workbook as HyperlinkExample.xlsx, and displays the full file location.
// Common Searches: Aspose.Cells C# add hyperlink to PDF on a Windows network share (UNC path) | How to insert a UNC path hyperlink into an Excel cell using Aspose.Cells library | C# create Excel file with clickable link to external PDF stored on \\ServerName\SharedFolder
// Tags: Aspose.Cells create hyperlink with UNC address | C# generate Excel link to PDF on shared network | Hyperlinks.Add usage for external PDF link | Save workbook containing external PDF link using Aspose.Cells | Excel cell hyperlink to PDF in shared directory

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, writes "Open PDF" into cell B2, adds a hyperlink that points to a PDF located on a UNC network share, saves the file as HyperlinkExample.xlsx, and outputs the full path of the saved workbook.
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

            // Set the display text for the hyperlink
            Cell targetCell = sheet.Cells["B2"];
            targetCell.PutValue("Open PDF");

            // UNC path to the external PDF on a network share
            string pdfPath = @"\\ServerName\SharedFolder\Document.pdf";

            // Add the hyperlink to cell B2 (zero‑based row/column indices)
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hyperlink address
            sheet.Hyperlinks.Add(1, 1, 1, 1, pdfPath);

            // Save the workbook
            string outputPath = "HyperlinkExample.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
