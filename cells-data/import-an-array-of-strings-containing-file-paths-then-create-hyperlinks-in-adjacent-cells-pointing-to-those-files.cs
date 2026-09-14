// Title: Import a string array of file paths into column A and generate "Open" hyperlinks in column B using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that reads a list of file paths, writes them to column A, and inserts a clickable "Open" hyperlink in column B for each row. | Show how to use Worksheet.Hyperlinks.Add to create file‑path hyperlinks next to values imported into an Excel worksheet with Aspose.Cells.
// Common Searches: aspnet load file path list into Excel and add hyperlinks with Aspose.Cells | c# Aspose.Cells create "Open" link in adjacent column for each file path | how to use Worksheet.Hyperlinks.Add after ImportArray in Aspose.Cells | generate Excel workbook containing file paths and open links using Aspose.Cells for .NET | add clickable link next to data column in Aspose.Cells workbook
// Tags: populate worksheet with file paths Aspose.Cells | add adjacent cell hyperlink Aspose.Cells | Worksheet.Hyperlinks.Add C# example | file path hyperlink column generation | hyperlink column creation Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new workbook, loads a string[] of file paths into column A, adds an "Open" hyperlink in column B for each path using Worksheet.Hyperlinks.Add, and saves the file as HyperlinksFromArray.xlsx.
class Program
{
    static void Main()
    {
        // Array of file paths to be imported
        string[] filePaths = new string[]
        {
            @"C:\Docs\file1.pdf",
            @"C:\Docs\file2.pdf",
            @"C:\Docs\file3.pdf"
        };

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Import the file paths vertically starting at cell A1 (row 0, column 0)
        cells.ImportArray(filePaths, 0, 0, true);

        // For each imported path, add a hyperlink in the adjacent column (B)
        for (int i = 0; i < filePaths.Length; i++)
        {
            // Convert row/column index to cell name (e.g., B1, B2, ...)
            string hyperlinkCell = CellsHelper.CellIndexToName(i, 1);

            // Add the hyperlink pointing to the file path
            worksheet.Hyperlinks.Add(hyperlinkCell, 1, 1, filePaths[i]);

            // Set display text for the hyperlink cell
            worksheet.Cells[hyperlinkCell].PutValue("Open");
        }

        // Save the workbook
        workbook.Save("HyperlinksFromArray.xlsx");
    }
}
